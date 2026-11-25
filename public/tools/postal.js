// 表格 · 郵便番号 ↔ 住所 批量补全（2026-11-25）：选一列邮编（或住所），右边一列写出住所（或邮编）。
// 数据是日本邮政官方的邮编数据（随安装包提供，可在侧栏「郵便番号」里更新）。只到町域；有几个候选时不乱猜，写出候选并标黄。
(function(root){'use strict';
var T=root.OfficeTools,L=T.lib,S=T.SheetPlus,cells=L.cells;
function n(v){return Number(v)||0;}
function text(v){return String(v==null?'':v);}
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
async function tool(body){var s=await L.localFetch('/api/session',{}).then(function(r){return r.json();});var r=await L.localFetch('/api/tools/postal',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(body)});var d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(d.error||M('本机服务没有完成，请重试。','ローカルサービスで処理できませんでした。もう一度お試しください。'));return d;}
var P=root.OfficePostal,looksZip=P.looksZip,render=P.render;
async function fill(c){var a=c.a,o=c.options||{},r=cells(a);if(n(r.Columns.Count)!==1)throw Error(M('请只选一列邮编或住所。','郵便番号か住所の列を 1 列だけ選んでください。'));S.limit(r,20000);
 var v=S.grid(r),vals=v.map(function(row,i){var x=row[0];if(typeof x==='number')x=String(x).padStart(7,'0');return text(x).trim();});var filled=vals.filter(Boolean);if(!filled.length)throw Error(M('选区是空的。','選択範囲が空です。'));
 var dir=o.dir==='zip2addr'||o.dir==='addr2zip'?o.dir:(filled.filter(looksZip).length>=filled.length/2?'zip2addr':'addr2zip');
 var out=S.rightColumn(a,r,1);
 var res=await tool({action:dir==='zip2addr'?'zip':'address',items:vals.map(function(x){return x||'';})});
 var rows=[],kinds=[],count={ok:0,multi:0,city:0,none:0,format:0};res.results.forEach(function(x,i){if(!vals[i]){rows.push(['']);kinds.push('');return;}var t=render(x,dir,o.fmt);count[t.kind]++;rows.push([t.t]);kinds.push(t.kind);});
 if(dir==='addr2zip')try{out.NumberFormat='@';}catch(ignore){}
 S.put(out,rows);try{out.Columns.Item(1).ColumnWidth=dir==='zip2addr'?34:12;}catch(ignore){}
 // 有几个候选：黄色；只到市区町村：浅橙；找不到：浅红
 kinds.forEach(function(k,i){var col=k==='multi'?65535:k==='city'?10079487:k==='none'||k==='format'?13421823:0;if(col)try{out.Cells.Item(i+1,1).Interior.Color=col;}catch(ignore){}});
 return (dir==='zip2addr'?M('已在右边一列写出住所：','右隣の列に住所を書きました：'):M('已在右边一列写出邮编：','右隣の列に郵便番号を書きました：'))+M('确定 '+count.ok+' 个','確定 '+count.ok+' 件')+(count.multi?M('，有几个候选 '+count.multi+' 个（黄色，请自己选）','、候補が複数 '+count.multi+' 件（黄色。選んでください）'):'')+(count.city?M('，只认到市区町村 '+count.city+' 个（浅橙色）','、市区町村までしか分からないもの '+count.city+' 件（薄いオレンジ）'):'')+(count.none+count.format?M('，没找到 '+(count.none+count.format)+' 个（浅红色）','、見つからないもの '+(count.none+count.format)+' 件（薄い赤）'):'')+M('。\n数据日期：','。\nデータの日付：')+(res.date||'?')+M('；只到町域，丁目和番地请自己补。','。町域までのため、丁目・番地は補ってください。');}
T.register({id:'postalFill',label:'邮编 ↔ 住所',desc:'日本邮编和住所互查，一列批量补全',icon:'postal',group:'sheetdata',undo:false,ownConfirm:['et'],keywords:'郵便番号 邮编 住所 地址 自動入力 補完 〒 zip postal',
 fields:[{key:'dir',label:'方向',options:[{value:'auto',label:'自动判断'},{value:'zip2addr',label:'邮编 → 住所'},{value:'addr2zip',label:'住所 → 邮编'}],value:'auto'},{key:'fmt',label:'住所写到哪一级',options:[{value:'full',label:'都道府县 + 市区町村 + 町域'},{value:'town',label:'只写町域'}],value:'full'}],
 hosts:{et:fill}});
})(window);
