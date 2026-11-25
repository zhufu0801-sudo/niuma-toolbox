// 办公工具箱 · 数据脱敏引擎（2026-11-25）：认号码、认词表里的词，算出替换成什么。
// 不猜姓名和地址：只认有固定格式的号码（身份证、银行卡、マイナンバー、手机号、邮箱）和客户自己填的词。
// 侧栏、本机服务（卡片 / 粘贴栏自动脱敏）和测试共用这一份；不碰文档，不联网。
(function(root){'use strict';
var TYPES=[
 {type:'WORD',label:'词表里的词',labelJa:'単語リストの語'},
 {type:'ID',label:'身份证号',labelJa:'身分証番号'},
 {type:'CARD',label:'银行卡号',labelJa:'カード番号'},
 {type:'MYNUMBER',label:'マイナンバー',labelJa:'マイナンバー'},
 {type:'PHONE',label:'手机号',labelJa:'携帯電話番号'},
 {type:'EMAIL',label:'邮箱',labelJa:'メールアドレス'}
];
var CODE={zh:{ID:'身份证',CARD:'银行卡',MYNUMBER:'个人编号',PHONE:'电话',EMAIL:'邮箱',WORD:'名称'},ja:{ID:'身分証',CARD:'カード',MYNUMBER:'マイナンバー',PHONE:'電話',EMAIL:'メール',WORD:'名称'}};
function luhn(d){var s=0,dbl=false;for(var i=d.length-1;i>=0;i--){var n=d.charCodeAt(i)-48;if(dbl){n*=2;if(n>9)n-=9;}s+=n;dbl=!dbl;}return s%10===0;}
function myNumberOk(d){if(!/^\d{12}$/.test(d))return false;var s=0;for(var n=1;n<=11;n++){var p=Number(d[11-n]),q=n<=6?n+1:n-5;s+=p*q;}var r=s%11,c=r<=1?0:11-r;return c===Number(d[11]);}
function idCnOk(id){var w=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2],c='10X98765432',s=0;for(var i=0;i<17;i++)s+=Number(id[i])*w[i];return c[s%11]===id[17].toUpperCase();}
// 顺序有讲究：先认长号码（身份证、银行卡、个人编号），再认手机号。和 lib/privacy.mjs 的规则一致。
var DETECTORS=[
 {type:'ID',re:/(?<![0-9A-Za-z])[1-9]\d{5}(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx](?![0-9A-Za-z])/g,ok:idCnOk},
 {type:'CARD',re:/(?<![\d-])\d{4}(?:[ -]?\d{4}){2,3}(?:[ -]?\d{1,3})?(?![\d-])/g,ok:function(m){var d=m.replace(/\D/g,'');return d.length>=15&&d.length<=19&&luhn(d);}},
 {type:'MYNUMBER',re:/(?<![\d-])\d{4}[ -]?\d{4}[ -]?\d{4}(?![\d-])/g,ok:function(m){return myNumberOk(m.replace(/\D/g,''));}},
 {type:'PHONE',re:/(?<![\d+])(?:\+?86[ -]?)?1[3-9]\d[ -]?\d{4}[ -]?\d{4}(?!\d)|(?<![\d+])(?:\+81[ -]?|0)[789]0[ -]?\d{4}[ -]?\d{4}(?!\d)/g},
 {type:'EMAIL',re:/[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}/g}
];
var MAX_WORDS=500;
// 词表：一行一个；“原文=代号”可以自己定代号；# 开头是备注。返回按长度从长到短排好的列表（长词优先，“北京某某科技有限公司”先于“北京”）。
function parseWords(text){var seen={},out=[],bad=[];String(text||'').split(/\r?\n/).forEach(function(line){var s=line.replace(/^\s+|\s+$/g,'');if(!s||s.charAt(0)==='#')return;var i=s.search(/[=＝]/),w=i>0?s.slice(0,i).trim():s,code=i>0?s.slice(i+1).trim():'';
  if(!w){bad.push(line);return;}if(w.length>100){bad.push(w.slice(0,20)+'…');return;}if(/[\r\n\t]/.test(w)){bad.push(w);return;}if(seen[w])return;seen[w]=1;out.push({word:w,code:code.slice(0,40)});});
 if(out.length>MAX_WORDS)throw Error('词表最多 '+MAX_WORDS+' 个词，请分成几套方案。');
 out.sort(function(a,b){return b.word.length-a.word.length;});return {words:out,bad:bad};}
function escapeRe(s){return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}
// 找出文字里要脱敏的地方。opts.types：要认哪些号码（不给 = 全部）；opts.words：parseWords 的结果。
// 返回不重叠、按位置排好的 [{start,end,value,type}]。词表优先，其次按号码规则的顺序。
function find(text,opts){text=String(text==null?'':text);opts=opts||{};var types=opts.types||DETECTORS.map(function(d){return d.type;}),words=opts.words||[],hits=[],taken=[];
 function free(s,e){for(var i=0;i<taken.length;i++)if(s<taken[i][1]&&e>taken[i][0])return false;return true;}
 function add(s,e,type){if(e<=s||!free(s,e))return;taken.push([s,e]);hits.push({start:s,end:e,value:text.slice(s,e),type:type});}
 if(words.length&&text){var re=new RegExp(words.map(function(w){return escapeRe(w.word);}).join('|'),'g'),m;while((m=re.exec(text))){if(!m[0]){re.lastIndex++;continue;}add(m.index,m.index+m[0].length,'WORD');}}
 DETECTORS.forEach(function(d){if(types.indexOf(d.type)<0)return;d.re.lastIndex=0;var m;while((m=d.re.exec(text))){if(d.ok&&!d.ok(m[0]))continue;add(m.index,m.index+m[0].length,d.type);}});
 hits.sort(function(a,b){return a.start-b.start;});return hits;}
function chars(s){return Array.from(String(s));}
// 只把数字换成 *，保留前 front 位、后 back 位数字和中间的分隔符（138-****-5678）。
function maskDigits(v,front,back){var total=(v.match(/\d/g)||[]).length,n=0;return v.replace(/\d/g,function(d){n++;return n<=front||n>total-back?d:'*';});}
function star(value,type){var v=String(value);
 if(type==='PHONE'){var plus=v.match(/^(\+?86[ -]?|\+81[ -]?)/),pre=plus?plus[0]:'',rest=v.slice(pre.length);return pre+maskDigits(rest,3,4);}
 if(type==='ID')return v.slice(0,3)+v.slice(3,-4).replace(/[\dXx]/g,'*')+v.slice(-4);
 if(type==='CARD'||type==='MYNUMBER')return maskDigits(v,0,4);
 if(type==='EMAIL'){var at=v.indexOf('@');return v.charAt(0)+'***'+v.slice(at);}
 var c=chars(v);if(c.length<=1)return '*';return c[0]+new Array(c.length).join('*');}
function block(value){return new Array(chars(value).length+1).join('■');}
function letters(n){var s='';n++;while(n>0){var r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}
// 一次脱敏的“会话”：同一个原文始终换成同一个代号；记下每个替换，方便存对照表。
// mode：star = 中间打码；block = 整段遮住；code = 换成代号。lang：zh / ja（只影响代号的名字）。
function session(opts){opts=opts||{};var mode=['star','block','code'].indexOf(opts.mode)>=0?opts.mode:'star',names=CODE[opts.lang==='ja'?'ja':'zh'],custom=Object.create(null),used=Object.create(null),byValue=Object.create(null),count=Object.create(null),pairs=[];
 (opts.words||[]).forEach(function(w){if(w.code)custom[w.word]=w.code;});
 // 接着上一次的代号编：同一份文档里再脱敏一次，不会出现两个【电话1】。
 (opts.prior||[]).forEach(function(p){if(p&&p.original!=null&&p.masked!=null&&!byValue[p.original]){byValue[p.original]=p.masked;used[p.masked]=1;if(p.type)count[p.type]=Math.max(count[p.type]||0,Number(p.n)||0);}});
 function code(value,type){if(type==='WORD'&&custom[value]){var c=custom[value],k=c,i=2;while(used[k])k=c+i++;return k;}
  var token;do{count[type]=(count[type]||0)+1;token='【'+names[type]+(type==='WORD'?letters(count[type]-1):count[type])+'】';}while(used[token]);return token;}
 function replace(value,type){var key=String(value);if(byValue[key]!=null)return byValue[key];var out=mode==='code'?code(key,type):mode==='block'?block(key):star(key,type);byValue[key]=out;used[out]=1;pairs.push({original:key,masked:out,type:type,n:count[type]||0});return out;}
 return {mode:mode,replace:replace,pairs:function(){return pairs.slice();}};}
// 把一段文字按 hits 换掉，返回新文字和每处的 [masked, original]（按出现顺序，还原时按顺序对回去）。
function apply(text,hits,sess){text=String(text==null?'':text);var out='',pos=0,seq=[];hits.slice().sort(function(a,b){return a.start-b.start;}).forEach(function(h){var m=sess.replace(h.value,h.type);out+=text.slice(pos,h.start)+m;pos=h.end;seq.push([m,h.value,h.type]);});return {text:out+text.slice(pos),seq:seq};}
function maskText(text,sess,opts){var hits=find(text,opts);return apply(text,hits,sess);}
// 还原一段文字：seq 是脱敏时记下的 [masked, original]，按出现顺序一处处换回。
// 返回 {text, done, missing}；找不到的（文字后来被改过）记在 missing 里，不乱换。
function restore(text,seq){var s=String(text==null?'':text),out='',pos=0,done=0,missing=[];(seq||[]).forEach(function(p){var i=s.indexOf(p[0],pos);if(i<0){missing.push(p);return;}out+=s.slice(pos,i)+p[1];pos=i+p[0].length;done++;});return {text:out+s.slice(pos),done:done,missing:missing};}
// 对照表：同一对原文 / 替换结果只列一行，带次数。
function table(seqs){var map={},list=[];(seqs||[]).forEach(function(seq){(seq||[]).forEach(function(p){var k=p[1]+'\u0000'+p[0];if(!map[k]){map[k]={original:p[1],masked:p[0],type:p[2]||'',count:0};list.push(map[k]);}map[k].count++;});});return list;}
// Restore all code occurrences in one pass; newest mappings go first. Never cascade
// replacements into restored source text. Normal prose in brackets stays untouched.
function restoreCodes(text,pairs){
 var known=Object.create(null),literal=Object.create(null),unknown=[],done=0;
 function key(s){return String(s).replace(/^[【\[［]|[】\]］]$/g,'').replace(/[\s\u3000]+/g,'');}
 (pairs||[]).forEach(function(p){if(!p||typeof p.masked!=='string'||!p.masked)return;var k=key(p.masked);if(!Object.prototype.hasOwnProperty.call(known,k))known[k]=String(p.original);if(!Object.prototype.hasOwnProperty.call(literal,p.masked))literal[p.masked]=String(p.original);});
 var exact=Object.keys(literal).sort(function(a,b){return b.length-a.length;}).map(escapeRe);
 var re=new RegExp('[【\\[［][^【】\\[\\]［］\\r\\n]{1,100}[】\\]］]'+(exact.length?'|'+exact.join('|'):''),'g');
 var result=String(text==null?'':text).replace(re,function(token){var k=key(token);if(Object.prototype.hasOwnProperty.call(literal,token)){done++;return literal[token];}if(Object.prototype.hasOwnProperty.call(known,k)){done++;return known[k];}if(/^(?:名称|电话|電話|邮箱|メール|身份证|身分証|银行卡|カード|个人编号|マイナンバー)[A-Za-z0-9]+$/.test(k)&&unknown.indexOf(token)<0)unknown.push(token);return token;});
 return {text:result,done:done,unknown:unknown};
}
function label(type,ja){var t=TYPES.find(function(x){return x.type===type;});return t?(ja?t.labelJa:t.label):type;}
var api={TYPES:TYPES,parseWords:parseWords,find:find,star:star,block:block,session:session,apply:apply,maskText:maskText,restore:restore,restoreCodes:restoreCodes,table:table,label:label,MAX_WORDS:MAX_WORDS};
root.MaskEngine=api;if(typeof module!=='undefined'&&module.exports)module.exports=api;
})(typeof window!=='undefined'?window:globalThis);
