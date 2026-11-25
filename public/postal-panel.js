// 办公工具箱 · 郵便番号 ↔ 住所 侧栏（2026-11-25）。在侧栏“文件”页（Office 版在工具列表下面）。
// 输入（或用选中的文字）邮编或住所，列出结果，点“插入”放到光标处；有几个候选时全部列出，让客户自己选。
// 数据：日本邮政官方邮编数据，随安装包提供；可以“更新数据”（从官网下载）或导入自己下载的 utf_ken_all.zip。
(function(root){'use strict';
var box=null,busy=false,last=[];
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
function tr(m){try{return ja()&&root.OfficeI18n&&root.OfficeI18n.t?root.OfficeI18n.t(m):m;}catch(e){return m;}}
function el(tag,cls,text){var e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;}
function $(id){return document.getElementById(id);}
function say(m,bad){var s=$('pcStatus');if(!s)return;s.textContent=m||'';s.className='mk-status'+(bad?' bad':'');}
function wps(){var T=root.OfficeTools;return T&&T.app&&T.app()?T:null;}
async function post(body){var f=root.fetch.bind(root),s=await f('/api/session').then(function(r){if(!r.ok)throw Error(M('本机服务没有连接。','ローカルサービスに接続できません。'));return r.json();});
 var r=await f('/api/tools/postal',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(body)}),d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(tr(d.error||M('没有完成，请重试。','完了できませんでした。もう一度お試しください。')));return d;}
function task(fn){if(busy)return;busy=true;box.querySelectorAll('button,input,select').forEach(function(b){b.disabled=true;});Promise.resolve().then(fn).catch(function(e){say(tr(String(e&&e.message||e)),true);}).then(function(){busy=false;box.querySelectorAll('button,input,select').forEach(function(b){b.disabled=false;});});}
function isZip(s){return /^〒?\s*\d{3}-?\d{4}$/.test(String(s||'').normalize('NFKC').trim());}
// ---- 选中的文字 / 插入 ----
async function selection(){var T=wps();if(T){var a=T.app(),h=T.host();try{if(h==='et')return String(a.ActiveCell.Text||'');if(h==='wpp')return String(a.ActiveWindow.Selection.TextRange.Text||'');return String(a.Selection.Text||'').replace(/[\r\u0007]+$/,'');}catch(e){return '';}}
 try{if(root.Word&&/Word/i.test(root.Office.context.host))return await root.Word.run(async function(c){var s=c.document.getSelection();s.load('text');await c.sync();return s.text;});
  if(root.Excel)return await root.Excel.run(async function(c){var r=c.workbook.getSelectedRange();r.load('text');await c.sync();return String(r.text[0][0]||'');});}catch(ignore){}return '';}
async function insert(t,isZipText){var T=wps();if(T){var a=T.app(),h=T.host();
  if(h==='et'){var c=a.ActiveCell;if(isZipText)try{c.NumberFormat='@';}catch(ignore){}c.Value2=t;return;}
  if(h==='wpp'){var sel=a.ActiveWindow.Selection;if(Number(sel.Type)<2)throw Error(M('请先点进一个文本框。','先にテキストボックス内をクリックしてください。'));sel.TextRange.Text=t;return;}
  a.Selection.TypeText(t);return;}
 if(root.Word&&/Word/i.test(root.Office.context.host))return root.Word.run(async function(c){c.document.getSelection().insertText(t,'Replace');await c.sync();});
 if(root.Excel)return root.Excel.run(async function(c){var r=c.workbook.getSelectedRange().getCell(0,0);if(isZipText)r.numberFormat=[['@']];r.values=[[t]];await c.sync();});
 throw Error(M('PowerPoint 版请复制后自己粘贴。','PowerPoint 版ではコピーして貼り付けてください。'));}
function fmtText(x){var f=$('pcFmt').value,addr=x.address;if(f==='zip')return '〒'+x.zip;if(f==='zipaddr')return '〒'+x.zip+' '+addr;return addr;}
// ---- 查询 ----
async function query(q){q=String(q||'').trim();if(!q)throw Error(M('请输入邮编或住所。','郵便番号か住所を入力してください。'));var zip=isZip(q),d=await post({action:zip?'zip':'address',items:[q]}),r=d.results[0];last=r.candidates||[];
 var list=$('pcList');list.textContent='';
 if(!r.ok){say(r.reason==='format'?M('邮编要 7 位数字。','郵便番号は 7 桁の数字です。'):r.reason==='city'?M('认不出市区町村。请从都道府县或市区町村开始写。','市区町村が分かりません。都道府県か市区町村から入力してください。'):M('没有找到。','見つかりませんでした。'),true);return;}
 last.forEach(function(x){var row=el('div','pc-row'),t=el('div','pc-text');t.append(el('b','','〒'+x.zip+' '),document.createTextNode(x.address));if(x.note)t.append(el('span','pc-note','（'+x.note+'）'));if(x.level==='city')t.append(el('span','pc-note',M(' · 只认到市区町村',' · 市区町村まで')));
  var b=el('button','',M('插入','挿入'));b.type='button';b.onclick=function(){task(async function(){var s=fmtText(x);await insert($('pcFmt').value==='zip'?s.replace(/^〒/,''):s,$('pcFmt').value==='zip');say(M('已插入：','挿入しました：')+s);});};row.append(t,b);list.append(row);});
 say((last.length>1?M('有 '+last.length+' 个候选，请选一个插入。','候補が '+last.length+' 件あります。1 つ選んで挿入してください。'):M('找到了。','見つかりました。'))+M('数据日期：','データの日付：')+(d.date||'?')+M('（只到町域）','（町域まで）'));}
async function status(){var s=await post({action:'status'}),line=$('pcData');if(!s.ok){line.textContent=tr(s.error);line.className='mk-note';return;}line.className='mk-meta';line.textContent=M('数据日期 '+s.date+'，共 '+s.rows+' 条（'+(s.from==='user'?'已更新':'安装包自带')+'）。日本邮政每月更新一次。','データの日付 '+s.date+'、'+s.rows+' 件（'+(s.from==='user'?'更新済み':'インストール時のもの')+'）。日本郵便は毎月更新しています。');}
function pickZip(){var T=wps();if(!T)return '';try{var dlg=T.app().FileDialog(3);dlg.Title=M('选择 utf_ken_all.zip','utf_ken_all.zip を選択');dlg.AllowMultiSelect=false;try{dlg.Filters.Clear();dlg.Filters.Add('zip / csv','*.zip;*.csv');}catch(ignore){}if(dlg.Show()!==-1)return null;return String(dlg.SelectedItems.Item(1));}catch(e){return '';}}
function build(){var host=document.querySelector('section[data-panel="files"]'),office=!host;if(office)host=$('toolSection');if(!host)return;var old=$('postalSection');if(old)old.remove();
 box=el('section','tool-card mask pc');box.id='postalSection';box.setAttribute('translate','no');if(!office){box.dataset.panel='files';box.hidden=host.hidden;}
 var head=el('div','card-heading'),icon=el('span','section-icon teal'),img=el('img');img.src=(office?'../':'')+'icons/postal.svg';img.alt='';icon.append(img);var h=el('div');h.append(el('h2','',M('郵便番号 ↔ 住所','郵便番号 ↔ 住所')),el('p','',M('日本的邮编和住所互查，结果可以直接插入；一整列请用表格里的「邮编 ↔ 住所」','日本の郵便番号と住所を相互に検索して挿入。列ごとの一括入力は表計算の「郵便番号 ↔ 住所」で')));head.append(icon,h);
 var row=el('div','pc-input'),inp=el('input');inp.id='pcQuery';inp.type='text';inp.placeholder=M('例如 100-0001 或 東京都千代田区千代田','例：100-0001 または 東京都千代田区千代田');inp.onkeydown=function(e){if(e.key==='Enter')task(function(){return query(inp.value);});};
 var go=el('button','primary',M('查询','検索'));go.type='button';go.onclick=function(){task(function(){return query(inp.value);});};row.append(inp,go);
 var bar=el('div','pc-bar'),useSel=el('button','',M('用选中的文字查','選択中の文字で検索'));useSel.type='button';useSel.onclick=function(){task(async function(){var s=(await selection()).trim();if(!s)throw Error(M('没有选中文字。','文字が選択されていません。'));inp.value=s;await query(s);});};
 var fmt=el('select');fmt.id='pcFmt';[['addr',M('插入住所','住所を挿入')],['zipaddr',M('插入 〒邮编 + 住所','〒番号 + 住所を挿入')],['zip',M('只插入邮编','郵便番号だけ挿入')]].forEach(function(x){var o=el('option','',x[1]);o.value=x[0];fmt.append(o);});bar.append(useSel,fmt);
 var list=el('div','pc-list');list.id='pcList';var st=el('p','mk-status');st.id='pcStatus';st.setAttribute('role','status');
 var data=el('p','mk-meta');data.id='pcData';var up=el('div','pc-bar'),upd=el('button','',M('更新数据','データを更新')),imp=el('button','',M('导入 zip…','zip を取り込む…'));upd.type=imp.type='button';
 upd.onclick=function(){task(async function(){say(M('正在从日本邮政网站下载（约 2MB）…','日本郵便のサイトからダウンロードしています（約 2MB）…'));var d=await post({action:'update'});await status();say(M('已更新：','更新しました：')+d.info.rows+M(' 条，数据日期 ',' 件、データの日付 ')+d.info.date+'。');});};
 imp.onclick=function(){task(async function(){var p=pickZip();if(p===null)return;if(!p){p=($('pcPath').value||'').trim();$('pcPath').hidden=false;if(!p)throw Error(M('这里打不开选择窗口，请把 utf_ken_all.zip 的完整位置粘贴到下面的框里，再点一次“导入 zip…”。','ここではファイル選択画面を開けません。utf_ken_all.zip の場所を下の欄に貼り付けて、もう一度「zip を取り込む…」を押してください。'));}var d=await post({action:'import',path:p});await status();say(M('已导入：','取り込みました：')+d.info.rows+M(' 条。',' 件。'));});};
 var path=el('input');path.id='pcPath';path.type='text';path.hidden=true;path.placeholder='C:\\Users\\…\\utf_ken_all.zip';up.append(upd,imp);
 box.append(head,row,bar,st,list,data,up,path);
 var anchor=$('macroCheckSection')||$('maskSection');if(anchor)anchor.after(box);else if(office)host.after(box);else host.parentNode.insertBefore(box,host);
 task(status);}
function start(){build();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start);else setTimeout(start,0);
try{if(root.OfficeI18n&&root.OfficeI18n.onChange)root.OfficeI18n.onChange(function(){build();});}catch(ignore){}
root.PostalPanel={query:function(q){return task(function(){return query(q);});}};
})(window);
