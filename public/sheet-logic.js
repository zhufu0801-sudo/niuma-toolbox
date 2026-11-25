// 办公工具箱 · 表格整理的纯计算部分（不碰表格对象），WPS 表格版和 Excel 版共用，便于测试。
(function(root){
  'use strict';
  function empty(v){return v===undefined||v===null||v==='';}
  var FW='０１２３４５６７８９．，－＋％';
  function half(s){return String(s).replace(/[０-９．，－＋％]/g,function(ch){return '0123456789.,-+%'[FW.indexOf(ch)];});}
  function serial(y,m,d){if(m<1||m>12||d<1||d>31||y<1900||y>9999)return null;var t=Date.UTC(y,m-1,d),dt=new Date(t);if(dt.getUTCMonth()!==m-1)return null;return Math.round(t/86400000)+25569;}
  // 文字 → 数字 / 日期。0 开头的编号、超过 15 位的数字（身份证、卡号）返回 null，保留文字。
  function parseText(s){var t=half(String(s)).replace(/[\s ]/g,'');if(!t)return null;var m;
    if((m=t.match(/^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?$/))){var d=serial(+m[1],+m[2],+m[3]);return d==null?null:{v:d,date:true};}
    if(/^0\d/.test(t))return null;
    var neg=/^\(.*\)$/.test(t);if(neg)t=t.slice(1,-1);var pct=/%$/.test(t);if(pct)t=t.slice(0,-1);t=t.replace(/^[¥￥$]/,'');
    if(!/^[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(t))return null;if(t.replace(/[^\d]/g,'').length>15)return null;
    var v=Number(t.replace(/,/g,''));if(!isFinite(v))return null;if(neg)v=-v;if(pct)v=v/100;return {v:v,pct:pct};}
  function squeeze(f){return String(f||'').replace(/\s/g,'');}
  function isSpot(f){return /^=OR\(ROW\(\)=\d+,COLUMN\(\)=\d+\)$/.test(squeeze(f));}
  function isBand(f){return /^=MOD\(ROW\(\)-\d+,2\)=1$/.test(squeeze(f));}
  function isDue(f){var s=squeeze(f);return /^=AND\(ISNUMBER\([A-Z]+\d+\),[A-Z]+\d+<TODAY\(\)\)$/.test(s)||/^=AND\(ISNUMBER\([A-Z]+\d+\),[A-Z]+\d+>=TODAY\(\),[A-Z]+\d+<=TODAY\(\)\+\d+\)$/.test(s);}
  // 新建工作表名、表头：日文界面用日文（判断“目录”“多表合并”时两种都认）。
  var NAMES={ja:{'工资条':'給与明細','拆分结果':'分割結果','合成一列':'1列にまとめる','转置':'行列入替','多表合并':'シート結合','目录':'目次','来源工作表':'元のシート','（工具箱生成，再次执行会更新）':'（ツールボックスが作成。再実行で更新）','（空白）':'（空白）','分栏打印':'段組み印刷','整理成表':'表に整理'}};
  function lang(){try{return root.OfficeI18n&&root.OfficeI18n.lang&&root.OfficeI18n.lang()==='ja'?'ja':'zh';}catch(e){return 'zh';}}
  function nm(zh){var l=lang();return l==='ja'&&NAMES.ja[zh]||zh;}
  function isIndex(name){return name==='目录'||name==='目次';}
  function isMerged(name){return /^(多表合并|シート結合)/.test(String(name));}
  function cleanName(s){s=String(s==null?'':s).replace(/[\\\/\?\*\[\]:']/g,'_').trim().slice(0,31);return s||'空白';}
  var EXTRACT={
    numbers:function(s){var m=half(s).match(/-?\d+(?:\.\d+)?/g);return m?m.join(' '):'';},
    chinese:function(s){var m=s.match(/[぀-ヿ㐀-鿿豈-﫿]+/g);return m?m.join(''):'';},
    letters:function(s){var m=s.match(/[A-Za-z]+(?:[ '\-][A-Za-z]+)*/g);return m?m.join(' '):'';},
    brackets:function(s){var t=s,prev;do{prev=t;t=t.replace(/[（(【\[][^（()）【】\[\]]*[）)】\]]/g,'');}while(t!==prev);return t.trim();}
  };
  function extract(kind,s){return EXTRACT[kind](String(s==null?'':s));}
  function key(v){return v==null?'':typeof v==='string'?v.trim():String(v);}
  // 两块区域：大小相同逐格比；不同按第一列对照。返回下标（从 0 开始）。
  function compare(va,vb){
    if(va.length===vb.length&&va[0].length===vb[0].length){var cells=[];for(var i=0;i<va.length;i++)for(var j=0;j<va[0].length;j++)if(key(va[i][j])!==key(vb[i][j]))cells.push([i,j]);return {same:true,cells:cells,onlyA:[],onlyB:[],diff:[]};}
    var ka={},kb={};va.forEach(function(r,i){ka[key(r[0])]=i;});vb.forEach(function(r,i){kb[key(r[0])]=i;});var onlyA=[],onlyB=[],diff=[];
    va.forEach(function(r,i){var k=key(r[0]);if(k==='')return;if(!(k in kb))onlyA.push(i);else{var rb=vb[kb[k]],w=Math.min(r.length,rb.length);for(var j=1;j<w;j++)if(key(r[j])!==key(rb[j])){diff.push([i,kb[k]]);break;}}});
    vb.forEach(function(r,i){var k=key(r[0]);if(k!==''&&!(k in ka))onlyB.push(i);});return {same:false,cells:[],onlyA:onlyA,onlyB:onlyB,diff:diff};}
  var SPLIT=/[、,，;；\/\n\r|｜]+/;
  function splitRows(head,rows,k){var out=head?[head]:[],split=0;rows.forEach(function(row){var parts=String(row[k]==null?'':row[k]).split(SPLIT).map(function(x){return x.trim();}).filter(Boolean);if(parts.length<2){out.push(row);return;}split++;parts.forEach(function(p){var c=row.slice();c[k]=p;out.push(c);});});return {rows:out,split:split};}
  function groupBy(values,k){var groups={},order=[];values.slice(1).forEach(function(row){var g=empty(row[k])?nm('（空白）'):String(row[k]);if(!groups[g]){groups[g]=[];order.push(g);}groups[g].push(row);});return {groups:groups,order:order};}
  function renamePlan(names,o){o=o||{};var find=String(o.find==null?'':o.find),rep=String(o.replace==null?'':o.replace),pre=String(o.prefix||''),suf=String(o.suffix||'');if(!find&&!pre&&!suf)throw Error('请填写要查找替换的文字，或要加的前缀、后缀。');
    var plan=[];names.forEach(function(x){var y=cleanName(pre+(find?x.split(find).join(rep):x)+suf);if(y!==x)plan.push([x,y]);});var finals={};names.forEach(function(x){finals[x.toLowerCase()]=x;});plan.forEach(function(p){delete finals[p[0].toLowerCase()];});
    plan.forEach(function(p){var low=p[1].toLowerCase();if(finals[low])throw Error('改名后“'+p[1]+'”会和已有工作表重名，请换一种改法。');finals[low]=p[1];});return plan;}
  // sheets: [{name, values}]（含表头的整张已用区域）
  function mergePlan(sheets){var head=null,parts=0,rows=[],mismatch=[];sheets.forEach(function(s){var v=s.values;if(!v.length||v.length===1&&v[0].every(empty))return;var h=v[0].map(function(x){return x==null?'':String(x).trim();});if(!head){head=h;rows.push(h.concat([nm('来源工作表')]));}if(h.join('\u0001')!==head.join('\u0001')){mismatch.push(s.name);return;}parts++;v.slice(1).forEach(function(row){if(row.some(function(x){return !empty(x);}))rows.push(row.concat([s.name]));});});return {parts:parts,rows:rows,mismatch:mismatch};}
  // ---- 表格体检 ----
  var HIDDEN=/[ ​-‍⁠﻿\t]/;
  function split(s){var out=[],depth=0,cur='',q=false;for(var i=0;i<s.length;i++){var ch=s[i];if(ch==='"')q=!q;if(!q&&ch==='(')depth++;if(!q&&ch===')')depth--;if(!q&&depth===0&&ch===','){out.push(cur);cur='';continue;}cur+=ch;}out.push(cur);return out;}
  function classify(cell){
    var out=[],v=cell.value,f=cell.formula,text=String(cell.text==null?'':cell.text);
    if(f){var up=f.toUpperCase(),m=up.match(/VLOOKUP\(([^()]|\([^()]*\))*\)/g);if(m)m.forEach(function(call){var args=split(call.slice(8,-1));if(args.length<4||/^(TRUE|1)$/.test(args[3].trim()))out.push('approx');});
      if(/^#(REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)/.test(text))out.push(/^#REF!/.test(text)?'ref':/^#N\/A/.test(text)?'na':'error');else if(/#REF!/.test(up))out.push('ref');}
    else if(typeof v==='string'){
      if(/^\s*=\s*([A-Za-z]{2,}\(|[A-Z]+\d)/.test(v))out.push('formulaText');else if(parseText(v))out.push('textNumber');
      if(/^[ 　]|[ 　]$/.test(v))out.push('spaces');if(HIDDEN.test(v))out.push('hidden');if(/^\d{16,}$/.test(v.trim()))out.push('longText');}
    else if(typeof v==='number'){if(Math.abs(v)>=1e15&&Number.isInteger(v))out.push('longNumber');if(/E\+\d+$/i.test(text))out.push('scientific');}
    return out;}
  var LABEL={
    textNumber:['看起来是数字、实际是文字','求和不算它、VLOOKUP 查不到。用“文本转数字”修复。'],
    spaces:['文字前后有空格','查找匹配不上。用“清理空格”修复。'],
    hidden:['含看不见的字符（不换行空格、零宽字符、Tab）','从网页或系统导出常见，匹配不上。用“查找替换”删掉，或重新输入。'],
    longText:['16 位以上的编号（文字）','保持文字格式是对的；不要转成数字，否则后几位会变 0。'],
    longNumber:['超过 15 位的数字','后几位已经变成 0（身份证、卡号常见），需要重新以文字格式输入。'],
    scientific:['显示成科学计数（如 1.23E+11）','把单元格格式改为“数值”或“文本”。'],
    formulaText:['公式显示成文字，不计算','单元格是“文本”格式。改为“常规”后双击单元格回车。'],
    approx:['VLOOKUP 用了近似匹配（第 4 个参数省略或为 TRUE）','没排序时会查出错值。第 4 个参数改为 FALSE 或 0。'],
    na:['#N/A 查不到','多半是类型不同（文字数字和真数字）或有空格，先看上面两项。'],
    ref:['#REF! 引用失效','引用的行列或工作表被删了，需要重新选引用范围。'],
    error:['其他公式错误（#VALUE!、#DIV/0! 等）','检查参与计算的单元格。']
  };
  var ORDER=['textNumber','spaces','hidden','formulaText','approx','na','ref','error','longNumber','scientific','longText'];
  function report(found,extra){var lines=[];ORDER.forEach(function(k){var x=found[k];if(!x||!x.count)return;lines.push('● '+LABEL[k][0]+'：'+x.count+' 处（'+x.where.slice(0,5).join('、')+(x.count>5?' 等':'')+'）\n  '+LABEL[k][1]);});(extra||[]).forEach(function(e){lines.push('● '+e);});return lines;}
  // ---- 打码：身份证、手机号、银行卡、邮箱、日本手机号；认不出的返回 null（不改）----
  function mask(v){var s=String(v==null?'':v).trim(),m;if(!s)return null;var t=half(s);
    if((m=t.match(/^(\d{6})(\d{8})(\d{3}[\dXx])$/)))return m[1]+'********'+m[3];
    if((m=t.match(/^(\d{6})(\d{6})(\d{3})$/)))return m[1]+'******'+m[3];
    if((m=t.match(/^(1[3-9]\d)(\d{4})(\d{4})$/)))return m[1]+'****'+m[3];
    if((m=t.match(/^(0[789]0)-?(\d{4})-?(\d{4})$/)))return m[1]+'-****-'+m[3];
    if((m=t.match(/^\d{12,19}$/)))return t.slice(0,4)+' **** **** '+t.slice(-4);
    if((m=s.match(/^([^@\s])[^@\s]*(@[^@\s]+\.[^@\s]+)$/)))return m[1]+'***'+m[2];
    return null;}
  // ---- 带单位求和：从“12元”“3.5kg”“-200 円”里取出数字 ----
  function unitSum(values){var sum=0,count=0,skipped=0,units={};values.forEach(function(v){if(empty(v))return;if(typeof v==='number'){sum+=v;count++;return;}var t=half(String(v)).replace(/,/g,''),m=t.match(/-?\d+(?:\.\d+)?/);if(!m){skipped++;return;}sum+=Number(m[0]);count++;var u=t.replace(m[0],'').replace(/[\s\d.+-]/g,'');if(u)units[u]=(units[u]||0)+1;});
    sum=Math.round(sum*1e8)/1e8;return {sum:sum,count:count,skipped:skipped,units:Object.keys(units)};}
  // ---- 细长表格分栏：表头 + 数据 → 每组并排，组之间空一列 ----
  function columnsFor(values,groups){var head=values[0],body=values.slice(1),w=head.length,per=Math.ceil(body.length/groups),out=[];
    for(var i=0;i<=per;i++){var row=[];for(var g=0;g<groups;g++){var src=i===0?head:body[g*per+i-1];for(var j=0;j<w;j++)row.push(src?src[j]:'');if(g<groups-1)row.push('');}out.push(row);}return {rows:out,per:per};}
  // ---- 一列整理成表：每 n 个非空单元格为一条记录 ----
  function listToTable(list,n){var items=list.filter(function(x){return !empty(x)&&String(x).trim()!=='';}),rows=[];for(var i=0;i<items.length;i+=n)rows.push(items.slice(i,i+n).concat(Array(Math.max(0,n-items.slice(i,i+n).length)).fill('')));return {rows:rows,rest:items.length%n};}
  root.OfficeSheetLogic={mask:mask,unitSum:unitSum,columnsFor:columnsFor,listToTable:listToTable,nm:nm,isIndex:isIndex,isMerged:isMerged,parseText:parseText,isSpot:isSpot,isBand:isBand,isDue:isDue,cleanName:cleanName,extract:extract,compare:compare,splitRows:splitRows,groupBy:groupBy,renamePlan:renamePlan,mergePlan:mergePlan,classify:classify,split:split,report:report,LABEL:LABEL,SPLIT:SPLIT};
})(typeof window!=='undefined'?window:globalThis);
