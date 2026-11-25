// 中日本地化小工具的计算部分（纯函数，WPS 版和 Office 版共用，不联网、不用 AI）。
(function(root){
 'use strict';
 var D='零壹贰叁肆伍陆柒捌玖',U=['','拾','佰','仟'],B=['','万','亿'];
 function four(x){var s='',zero=false;for(var i=3;i>=0;i--){var d=Math.floor(x/Math.pow(10,i))%10;if(d){if(zero)s+='零';s+=D[d]+U[i];zero=false;}else if(s)zero=true;}return s;}
 // 金额转人民币大写：1234.56 → 壹仟贰佰叁拾肆元伍角陆分；不能识别的返回 null
 function rmbUpper(v){
  if(v===null||v===undefined||v==='')return null;var n=typeof v==='number'?v:Number(String(v).replace(/[,，¥￥\s]|元$/g,''));
  if(!isFinite(n)||Math.abs(n)>=1e12)return null;
  var neg=n<0,fen=Math.round(Math.abs(n)*100),yuan=Math.floor(fen/100),jiao=Math.floor(fen%100/10),f=fen%10,int='';
  if(yuan){var parts=[yuan%10000,Math.floor(yuan/1e4)%1e4,Math.floor(yuan/1e8)],s='',need=false;
   for(var k=2;k>=0;k--){var p=parts[k];if(p){if(need||(s&&p<1000))s+='零';s+=four(p)+B[k];need=false;}else if(s)need=true;}int=s+'元';}
  var frac=jiao||f?(jiao?D[jiao]+'角':(int?'零':''))+(f?D[f]+'分':''):'整';
  if(!int&&!(jiao||f))return '零元整';
  return (neg?'负':'')+int+frac;
 }
 // 身份证号：18 位校验校验位；15 位旧号码按 19xx 年处理
 function idInfo(v){
  var id=String(v||'').trim().toUpperCase();
  if(/^\d{15}$/.test(id))return {birth:'19'+id.slice(6,8)+'-'+id.slice(8,10)+'-'+id.slice(10,12),gender:Number(id[14])%2?'男':'女',note:'15 位旧号码'};
  if(!/^\d{17}[\dX]$/.test(id))return {error:'不是 18 位身份证号'};
  var w=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2],c='10X98765432',s=0;for(var i=0;i<17;i++)s+=Number(id[i])*w[i];
  if(c[s%11]!==id[17])return {error:'校验位不对，请核对号码'};
  var y=Number(id.slice(6,10)),m=Number(id.slice(10,12)),d=Number(id.slice(12,14)),dt=new Date(Date.UTC(y,m-1,d));
  if(dt.getUTCFullYear()!==y||dt.getUTCMonth()!==m-1||dt.getUTCDate()!==d)return {error:'出生日期不存在'};
  return {birth:id.slice(6,10)+'-'+id.slice(10,12)+'-'+id.slice(12,14),gender:Number(id[16])%2?'男':'女'};
 }
 // 中国大陆手机号检查
 function phoneCheck(v){
  if(v===null||v===undefined||v==='')return null;var raw=String(v).replace(/[\s\-()（）]/g,'').replace(/^(\+?86)/,'');
  if(!/^\d+$/.test(raw))return '含有非数字字符';if(raw.length!==11)return '位数不对（'+raw.length+' 位）';if(!/^1[3-9]/.test(raw))return '号段不对';
  return '有效：'+raw.slice(0,3)+' '+raw.slice(3,7)+' '+raw.slice(7);
 }
 // 全角半角：英数字和常用符号变半角，半角片假名变全角（日本公文常用规则）。全角空格和中日文标点不动。
 // 中文全角逗号、冒号和中日文括号、感叹号、问号保持全角（它们是正常标点）
 var HALF_SYMBOLS='-.@/_%#&+=*';
 function zenHan(v,mode){
  if(typeof v!=='string')return null;var s=v;
  if(mode==='full')s=s.replace(/[A-Za-z0-9]/g,function(ch){return String.fromCharCode(ch.charCodeAt(0)+0xFEE0);});
  else s=s.replace(/[！-～]/g,function(ch){var h=String.fromCharCode(ch.charCodeAt(0)-0xFEE0);return /[A-Za-z0-9]/.test(h)||HALF_SYMBOLS.indexOf(h)>=0?h:ch;});
  s=s.replace(/[｡-ﾟ]+/g,function(run){return run.normalize('NFKC');});
  return s===v?null:s;
 }
 // 西历 ⇄ 和历
 var ERAS=[{n:'令和',a:'R',s:[2019,5,1]},{n:'平成',a:'H',s:[1989,1,8]},{n:'昭和',a:'S',s:[1926,12,25]},{n:'大正',a:'T',s:[1912,7,30]},{n:'明治',a:'M',s:[1868,10,23]}];
 function key(y,m,d){return y*10000+m*100+d;}
 function toDateParts(v){
  if(typeof v==='number'&&v>0&&v<200000){var dt=new Date(Math.round((v-25569)*86400000));return [dt.getUTCFullYear(),dt.getUTCMonth()+1,dt.getUTCDate()];}
  var m=String(v||'').trim().match(/^(\d{4})\s*[年\/.\-]\s*(\d{1,2})\s*[月\/.\-]\s*(\d{1,2})\s*日?$/);if(!m)return null;
  var y=+m[1],mo=+m[2],d=+m[3],t=new Date(Date.UTC(y,mo-1,d));return t.getUTCMonth()===mo-1&&t.getUTCDate()===d?[y,mo,d]:null;
 }
 function toWareki(v){
  var p=toDateParts(v);if(!p)return null;var k=key(p[0],p[1],p[2]);
  for(var i=0;i<ERAS.length;i++){var e=ERAS[i];if(k>=key(e.s[0],e.s[1],e.s[2])){var y=p[0]-e.s[0]+1;return e.n+(y===1?'元':y)+'年'+p[1]+'月'+p[2]+'日';}}
  return null;
 }
 // 返回 Excel 日期序列号（写回时设成日期格式）；不在年号范围内返回 {error}
 function toSeireki(v){
  if(typeof v!=='string')return null;var m=v.trim().match(/^(明治|大正|昭和|平成|令和|[MTSHRmtshr])\s*(元|\d{1,2})\s*[年.\/\-]\s*(\d{1,2})\s*[月.\/\-]\s*(\d{1,2})\s*日?$/);if(!m)return null;
  var e=ERAS.filter(function(x){return x.n===m[1]||x.a===m[1].toUpperCase();})[0],ey=m[2]==='元'?1:+m[2],y=e.s[0]+ey-1,mo=+m[3],d=+m[4],t=new Date(Date.UTC(y,mo-1,d));
  if(t.getUTCMonth()!==mo-1||t.getUTCDate()!==d)return {error:'日期不存在'};
  var k=key(y,mo,d),idx=ERAS.indexOf(e),next=ERAS[idx-1];
  if(k<key(e.s[0],e.s[1],e.s[2])||(next&&k>=key(next.s[0],next.s[1],next.s[2])))return {error:'不在'+e.n+'的范围内'};
  return {serial:Math.round(t.getTime()/86400000)+25569,text:y+'/'+mo+'/'+d};
 }
 // 邮编：日本 7 位统一成 123-4567；中国 6 位保持数字。其他返回 null（不改）
 function postal(v){
  if(v===null||v===undefined||v==='')return null;var raw=String(v).replace(/[０-９]/g,function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0);}).replace(/^〒\s*/,'').replace(/[\s\-ー－‐]/g,'');
  if(/^\d{7}$/.test(raw)){var out=raw.slice(0,3)+'-'+raw.slice(3);return out===String(v)?null:out;}
  if(/^\d{6}$/.test(raw))return raw===String(v)?null:raw;
  return null;
 }
 root.LocaleEngine={rmbUpper:rmbUpper,idInfo:idInfo,phoneCheck:phoneCheck,zenHan:zenHan,toWareki:toWareki,toSeireki:toSeireki,postal:postal,ERAS:ERAS};
})(typeof window!=='undefined'?window:globalThis);
