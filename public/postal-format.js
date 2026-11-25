// 办公工具箱 · 邮编批量补全的共用部分（2026-11-25）：WPS 表格（tools/postal.js）和 Office 版 Excel（office-batch2.js）都用。
(function(root){'use strict';
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
// 一格是邮编还是住所：7 位数字（可带 〒 和 -）就是邮编
function looksZip(s){return /^〒?\s*\d{3}-?\d{4}$/.test(String(s||'').normalize('NFKC').trim());}
// 把查询结果变成要写的文字和标记：ok / multi（有几个候选）/ city（只认到市区町村）/ none
function render(r,dir,fmt){if(!r||!r.ok)return {t:'',kind:r&&r.reason==='format'?'format':'none'};var c=r.candidates;
 if(dir==='zip2addr'){var list=c.map(function(x){return fmt==='town'?x.town||x.city:x.address;});var uniq=list.filter(function(v,i){return list.indexOf(v)===i;});if(uniq.length>1)return {t:M('候选：','候補：')+uniq.join(' / '),kind:'multi'};return {t:uniq[0],kind:'ok'};}
 var zips=c.map(function(x){return x.zip;}).filter(function(v,i,a){return a.indexOf(v)===i;});if(zips.length>1)return {t:M('候选：','候補：')+c.map(function(x){return x.zip+' '+x.town+(x.note?'（'+x.note+'）':'');}).join(' / '),kind:'multi'};
 return {t:zips[0],kind:c[0].level==='city'?'city':'ok'};}
root.OfficePostal={looksZip:looksZip,render:render};
})(typeof window!=='undefined'?window:globalThis);
