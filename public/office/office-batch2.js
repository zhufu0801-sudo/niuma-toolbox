// 办公工具箱 · Microsoft Office 版的第二批新功能（2026-11-24）。规则和 WPS 版一样；Office.js 做不到的给出操作步骤。
// 挂在 OfficeAddinTools 上：list 里多出这些工具，run 先找这里的实现。尚未在真实 Microsoft Office 中验证。
(function(root){'use strict';
var base=root.OfficeAddinTools;if(!base)return;
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
function need(c,msg){if(!c)throw Error(msg);}
function empty(v){return v===''||v===null||v===undefined;}
async function post(f,url,data){need(typeof f==='function',M('本机服务没有连接，请重新打开侧栏。','ローカルサービスに接続できません。サイドバーを開き直してください。'));var s=await f('/api/session').then(function(r){return r.json();}),r=await f(url,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(data)}),d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(d.error||M('操作失败。','処理できませんでした。'));return d;}
var CANCEL=function(){return M('已取消。','キャンセルしました。');};

// ---- Excel ----
var UNITS={wan:'0"."0,"万"',yi:'0"."00,,"亿"',qian:'#,##0,',baiwan:'#,##0,,',senyen:'#,##0,"千円"',general:'General'};
async function wanUnit(api,o){var u=UNITS[o.unit]?o.unit:'wan';return api.Excel.run(async function(c){var r=c.workbook.getSelectedRange();r.load('values,rowCount,columnCount');await c.sync();need(r.rowCount*r.columnCount<=200000,M('选区太大，请缩小范围。','選択範囲が大きすぎます。'));
  var nums=0;r.values.forEach(function(row){row.forEach(function(v){if(typeof v==='number')nums++;});});need(nums,M('选区里没有数字。','選択範囲に数値がありません。'));
  r.numberFormat=r.values.map(function(row){return row.map(function(){return UNITS[u];});});await c.sync();
  return u==='general'?M('已把选区恢复成普通数字格式。','選択範囲を標準の数値表示に戻しました。'):M('已把 '+nums+' 个数字按“'+{wan:'万',yi:'亿',qian:'千',baiwan:'百万',senyen:'千円'}[u]+'”显示。只改显示方式，数值不变，可以按 Ctrl+Z 撤销。',nums+' 個の数値の表示単位を変えました。値はそのままです。Ctrl+Z で戻せます。');});}
function rightBlank(c,r){var out=r.getOffsetRange(0,r.columnCount).getResizedRange(0,1-r.columnCount);out.load('values,address');return out;}
async function idCheck(api,o){var I=root.OfficeIdCheck;need(I,M('校验组件没有加载，请刷新侧栏。','チェック機能が読み込まれていません。サイドバーを更新してください。'));
  return api.Excel.run(async function(c){var r=c.workbook.getSelectedRange();r.load('values,text,rowCount,columnCount');await c.sync();need(r.columnCount===1,M('请只选一列号码。','番号の列を 1 列だけ選んでください。'));need(r.rowCount<=20000,M('选区太大，请缩小范围。','選択範囲が大きすぎます。'));
    var out=rightBlank(c,r);await c.sync();need(out.values.every(function(x){return empty(x[0]);}),M('结果要写在右边一列，但那里已有内容。请先插入空白列。','結果は右隣の列に書きますが、そこに内容があります。先に空の列を挿入してください。'));
    var good=0,bad=0,res=r.values.map(function(row,i){var v=row[0];if(empty(v))return [''];var msg;if(typeof v==='number'&&Math.abs(v)>=1e15)msg=M('✗ 这个号码按数字保存，超过 15 位的部分已经丢失。','✗ 数値として保存され、15 桁を超える部分が失われています。');else{var k=I.check(typeof v==='number'?String(r.text[i][0]||v):String(v),o.type==='auto'?'':o.type,{lang:ja()?'ja':'zh'});msg=k?k.msg:'';}if(/^✓/.test(msg))good++;else if(msg)bad++;return [msg];});
    need(good+bad,M('选区里没有号码。','選択範囲に番号がありません。'));out.values=res;out.format.columnWidth=260;await c.sync();
    return M('已在右边一列写出校验结果：通过 '+good+' 个，有问题 '+bad+' 个。只按编码规则检查校验位，不能证明号码真实存在。','右隣の列に結果を書きました：OK '+good+' 件、問題あり '+bad+' 件。規則とチェックデジットだけの確認で、実在するかは分かりません。');});}
function normMark(v){var t=String(v==null?'':v).trim();if(/^[√✓✔☑]$/.test(t))return '√';if(/^[○〇◯]$/.test(t))return '○';if(/^[×✗✘xX]$/.test(t))return '×';return t;}
async function attendanceStats(api){return api.Excel.run(async function(c){var r=c.workbook.getSelectedRange();r.load('values,rowCount,columnCount');await c.sync();var v=r.values;
  need(v.length>=2&&v[0].length>=3,M('请连同表头选中整张考勤表：第一行是日期，左边一列是姓名，中间是 √、休、假、迟 等记号。','見出しを含めて勤怠表全体を選んでください（1 行目が日付、左が氏名、間に √・休 などの記号）。'));
  var cols=v[0].length,nameCol=-1;for(var j=0;j<cols&&nameCol<0;j++){var k=0;for(var i=1;i<v.length;i++){var x=v[i][j];if(typeof x==='string'&&x.trim().length>=2&&!/^\d+$/.test(x.trim()))k++;}if(k>=Math.max(1,(v.length-1)/2))nameCol=j;}
  need(nameCol>=0,M('找不到姓名列。','氏名の列が見つかりません。'));var days=[];for(var j2=nameCol+1;j2<cols;j2++){var seen={},ok=true,any=false;for(var i2=1;i2<v.length;i2++){var y=v[i2][j2];if(empty(y))continue;any=true;var t=normMark(y);if(t.length>4){ok=false;break;}seen[t]=1;}if(ok&&any&&Object.keys(seen).length<=12)days.push(j2);}
  need(days.length,M('姓名右边没有找到记考勤的列。','氏名の右に勤怠の列が見つかりません。'));var people=[],by={},tot={};
  for(var i3=1;i3<v.length;i3++){var nm=String(v[i3][nameCol]||'').trim();if(!nm)continue;var p=by[nm]||(by[nm]={name:nm,c:{}});if(people.indexOf(p)<0)people.push(p);days.forEach(function(j){var m=normMark(v[i3][j]);if(!m)return;p.c[m]=(p.c[m]||0)+1;tot[m]=(tot[m]||0)+1;});}
  var marks=Object.keys(tot).sort(function(a,b){return tot[b]-tot[a];});need(marks.length,M('考勤列里都是空的。','勤怠の列がすべて空です。'));
  var J=ja(),out=[[J?'氏名':'姓名'].concat(marks).concat([J?'記録日数':'记录天数'])];people.forEach(function(p){var s=0,l=[p.name];marks.forEach(function(m){var n=p.c[m]||0;s+=n;l.push(n);});l.push(s);out.push(l);});
  var tl=[J?'合計':'合计'];marks.forEach(function(m){tl.push(tot[m]);});tl.push(marks.reduce(function(s,m){return s+tot[m];},0));out.push(tl);
  var sheets=c.workbook.worksheets;sheets.load('items/name');await c.sync();var names=sheets.items.map(function(s){return s.name;}),title=J?'勤怠集計':'考勤统计',name=title,k2=2;while(names.indexOf(name)>=0)name=title+'('+(k2++)+')';
  var ws=sheets.add(name),rg=ws.getRangeByIndexes(0,0,out.length,out[0].length);rg.values=out;rg.format.autofitColumns();ws.getRangeByIndexes(0,0,1,out[0].length).format.font.bold=true;ws.activate();await c.sync();
  return M('已在新工作表“'+name+'”里统计 '+people.length+' 人、'+days.length+' 天。原表不变。','新しいシート「'+name+'」に '+people.length+' 人・'+days.length+' 日分を集計しました。元の表は変更していません。');});}
async function codesBatch(api,o){var kind=o.kind==='barcode'?'barcode':'qr';return api.Excel.run(async function(c){var r=c.workbook.getSelectedRange();r.load('values,text,rowCount,columnCount');await c.sync();
  need(r.columnCount===1,M('请只选一列内容（网址、编号等）。','URL や番号の列を 1 列だけ選んでください。'));need(r.rowCount<=200,M('Office 版一次最多 200 个。','Office 版では一度に 200 個までです。'));
  var out=rightBlank(c,r);await c.sync();need(out.values.every(function(x){return empty(x[0]);}),M('右边一列已有内容，请先插入空白列。','右隣の列に内容があります。先に空の列を挿入してください。'));
  var items=r.values.map(function(row,i){return typeof row[0]==='number'?String(r.text[i][0]||row[0]):String(row[0]==null?'':row[0]);});need(items.some(function(x){return x.trim();}),M('选区是空的。','選択範囲が空です。'));
  if(!o.confirmed)return {confirm:M('将为选中的格子生成'+(kind==='qr'?'二维码':'条形码')+'，插到右边一列，并调整行高列宽。是否继续？','選んだセルの'+(kind==='qr'?'QR コード':'バーコード')+'を作り、右隣の列に貼り付けて行の高さと列の幅を調整します。続行しますか？')};
  var res=await post(o.fetch,'/api/tools/codes',{items:items,kind:kind,base64:true}),sh=c.workbook.worksheets.getActiveWorksheet().shapes;
  out.format.columnWidth=kind==='qr'?64:140;out.format.rowHeight=kind==='qr'?70:46;var cells=[];res.data.forEach(function(d,i){if(!d)return;var cell=out.getCell(i,0);cell.load('left,top,width,height');cells.push({cell:cell,d:d});});await c.sync();
  cells.forEach(function(x){x.img=sh.addImage(x.d);x.img.load('width,height');});await c.sync();
  cells.forEach(function(x){var img=x.img,W=x.cell.width-6,H=x.cell.height-6,ratio=kind==='qr'?1:img.width/img.height,w=Math.min(W,H*ratio),h=w/ratio;img.lockAspectRatio=false;img.width=w;img.height=h;img.left=x.cell.left+3+(W-w)/2;img.top=x.cell.top+3;img.placement='TwoCell';});await c.sync();
  return M('已插入 '+cells.length+' 个'+(kind==='qr'?'二维码':'条形码')+'，图片也存在：',(kind==='qr'?'QR コード':'バーコード')+'を '+cells.length+' 個貼り付けました。画像の保存先：')+res.folder+((res.errors||[]).length?M('\n没生成的（选区内行号）：','\n作成できなかったもの（選択範囲内の行番号）：')+res.errors.slice(0,5).map(function(e){return M('第 '+(e.index+1)+' 行','行 '+(e.index+1))+'（'+(root.OfficeI18n?root.OfficeI18n.t(e.error):e.error)+'）';}).join(M('；','、')):'');});}
async function excelPhrase(api,o,text){return api.Excel.run(async function(c){var cell=c.workbook.getActiveCell();cell.load('values');await c.sync();if(!empty(cell.values[0][0])&&!o.confirmed)return {confirm:M('选中的格子里已经有内容，要换成这句常用语吗？','選んだセルにはすでに内容があります。この定型文に置き換えますか？')};cell.values=[[text]];await c.sync();return M('已写入常用语。','定型文を入力しました。');});}

// ---- Word ----
var MARKS=['【办公工具箱】','【オフィスツールボックス】'];
async function proofCheck(api,o){var P=root.OfficeProof;need(P,M('检查组件没有加载，请刷新侧栏。','チェック機能が読み込まれていません。サイドバーを更新してください。'));var mode=['comment','fix','clear'].indexOf(o.mode)>=0?o.mode:'comment';
  return api.Word.run(async function(c){var body=c.document.body;
    if(mode==='clear'){var cs=body.getComments();cs.load('items/content');await c.sync();var gone=0;cs.items.forEach(function(x){if(MARKS.some(function(m){return String(x.content||'').indexOf(m)===0;})){x.delete();gone++;}});await c.sync();return gone?M('已删除工具加的 '+gone+' 条检查批注。','チェック用コメント '+gone+' 件を削除しました。'):M('没有找到工具加的检查批注。','ツールが付けたチェック用コメントはありません。');}
    var sel=c.document.getSelection();sel.load('text');await c.sync();var scope=String(sel.text||'').replace(/\s/g,'').length>1?sel:body,ps=scope.paragraphs;ps.load('items/text');await c.sync();need(ps.items.length<=5000,M('文档太长，请选中一部分再检查。','文書が長すぎます。一部を選択してからチェックしてください。'));
    var found=[];ps.items.forEach(function(p){var t=String(p.text||'');P.check(t,{lang:ja()?'ja':'zh'}).forEach(function(x){if(x.text.length>200)return;var nth=0,at=-1;while((at=t.indexOf(x.text,at+1))>=0&&at<x.start)nth++;found.push({p:p,x:x,nth:nth});});});
    if(!found.length)return M('没有发现常见错别字和标点问题。（只按规则检查，不能代替人工校对。）','よくある誤字や句読点の問題は見つかりませんでした。（ルールによる簡易チェックです。）');
    if(mode==='fix'&&!o.confirmed)return {confirm:M('发现 '+found.length+' 处可能的问题。能确定改法的直接改，其余加批注。可以按 Ctrl+Z 撤销。是否继续？','問題の可能性がある箇所が '+found.length+' 件あります。確実なものは直接直し、残りにはコメントを付けます。Ctrl+Z で戻せます。続行しますか？')};
    var searches=found.slice(0,300).map(function(f){var s=f.p.search(f.x.text,{matchCase:true,matchWildcards:false});s.load('items');return {f:f,s:s};});await c.sync();
    var fixed=0,noted=0,mk=MARKS[ja()?1:0];searches.forEach(function(q){var r=q.s.items[q.f.nth];if(!r)return;if(mode==='fix'&&q.f.x.fix!=null){r.insertText(q.f.x.fix,'Replace');fixed++;}else{r.insertComment(mk+q.f.x.msg);noted++;}});await c.sync();
    return mode==='fix'?M('已改正 '+fixed+' 处，'+noted+' 处加了批注。',fixed+' か所を直し、'+noted+' か所にコメントを付けました。'):M('发现 '+found.length+' 处可能的问题，已加 '+noted+' 条批注（以“'+mk+'”开头），看完后选“清除检查批注”可以一次删掉。','問題の可能性がある箇所 '+found.length+' 件に、コメントを '+noted+' 件付けました（「'+mk+'」で始まります）。確認後「チェック用コメントを削除」でまとめて消せます。');});}
async function sendCheck(api,o){return api.Word.run(async function(c){var d=c.document,body=d.body,cs=body.getComments(),tc=null,props=d.properties;cs.load('items');props.load('author,lastAuthor,company,manager');try{d.load('changeTrackingMode');}catch(e){}try{tc=body.getTrackedChanges();tc.load('items');}catch(e){tc=null;}
  try{await c.sync();}catch(e){tc=null;cs=body.getComments();cs.load('items');props.load('author,lastAuthor,company,manager');await c.sync();}
  var lines=[],ppl=[];if(cs.items.length)lines.push(M('批注 '+cs.items.length+' 条','コメント '+cs.items.length+' 件'));if(tc&&tc.items.length)lines.push(M('没有接受的修订 '+tc.items.length+' 处','未承諾の変更履歴 '+tc.items.length+' 件'));var tracking=d.changeTrackingMode&&d.changeTrackingMode!=='Off';if(tracking)lines.push(M('“修订”模式还开着','変更履歴の記録がオン'));
  ['author','lastAuthor','manager','company'].forEach(function(k){if(props[k])ppl.push(props[k]);});if(ppl.length)lines.push(M('文件属性里有：','プロパティ：')+ppl.join('、'));
  if(!lines.length)return M('检查完毕：没有批注和修订，文件属性里也没有作者信息。（Office 版不检查隐藏文字和高亮。）','チェック完了：コメントと変更履歴はなく、作成者情報もありません。（Office 版では隠し文字と蛍光ペンは調べません。）');
  if(!o.confirmed)return {confirm:M('发送前检查发现：\n· ','送付前チェックの結果：\n· ')+lines.join('\n· ')+M('\n\n要一键清理吗？（删除批注、接受修订、关闭修订模式、清空作者和单位）','\n\nまとめて片付けますか？（コメント削除・変更履歴の承諾・記録オフ・作成者と会社名の消去）')};
  cs.items.forEach(function(x){x.delete();});if(tc&&tc.items.length)tc.acceptAll();if(tracking)d.changeTrackingMode='Off';props.author='';props.company='';props.manager='';await c.sync();
  return M('已清理。保存后再发送。（Office 版不检查隐藏文字和高亮，需要时请在“文件 → 信息 → 检查问题”里检查。）','片付けました。保存してから送ってください。（隠し文字などは「ファイル → 情報 → 問題のチェック」で確認できます。）');});}
// 日文文体转换：本机服务算出每段要换的地方，这里按“原文片段 + 第几次出现”找到再替换。
async function jaStyle(api,o){var to=o.to==='polite'?'polite':'plain',plain=o.plain==='da'?'da':'dearu';
  return api.Word.run(async function(c){var sel=c.document.getSelection();sel.load('text');await c.sync();var scope=String(sel.text||'').replace(/\s/g,'').length>1?sel:c.document.body,ps=scope.paragraphs;ps.load('items/text');await c.sync();
    need(ps.items.length<=5000,M('文档太长，请选中一部分再转换。','文書が長すぎます。一部を選択してから変換してください。'));var texts=ps.items.map(function(p){return String(p.text||'');});
    need(texts.some(function(t){return /[぀-ヿ]/.test(t);}),M('选中的内容里没有日文。','選択範囲に日本語の文章がありません。'));
    var res=await post(o.fetch,'/api/tools/ja-style',{paragraphs:texts,to:to,plain:plain});if(!res.edits)return M('没有需要转换的句子。','変換する文は見つかりませんでした。');
    if(!o.confirmed)return {confirm:M('将改 '+res.edits+' 处句末的写法（「」里的引用不改），可以按 Ctrl+Z 撤销。是否继续？',res.edits+' か所の文末を変えます（「」内の引用は変えません）。Ctrl+Z で戻せます。続行しますか？')};
    var jobs=[];res.paragraphs.forEach(function(p,i){var t=texts[i];p.edits.forEach(function(e){var s=e.start,from=e.from,rep=e.text;if(!from){var back=Math.min(2,s);s-=back;from=t.slice(s,e.start);rep=from+rep;}if(!from||from.length>200)return;var nth=0,at=-1;while((at=t.indexOf(from,at+1))>=0&&at<s)nth++;var q=ps.items[i].search(from,{matchCase:true,matchWildcards:false});q.load('items');jobs.push({q:q,nth:nth,rep:rep});});});
    await c.sync();var done=0;jobs.reverse().forEach(function(j){var r=j.q.items[j.nth];if(!r)return;if(j.rep)r.insertText(j.rep,'Replace');else r.insertText('','Replace');done++;});await c.sync();
    return M('已转换 '+done+' 处。只改句末等处的词尾，建议再通读一遍。',done+' か所を変換しました。文末などの語尾だけを変えているので、念のため通して読んでください。');});}
async function wordPhrase(api,o,text){return api.Word.run(async function(c){c.document.getSelection().insertText(text,'End');await c.sync();return M('已插入常用语。','定型文を挿入しました。');});}
function pptPhrase(api,o,text){return new Promise(function(resolve,reject){try{api.Office.context.document.setSelectedDataAsync(text,function(r){if(r.status==='succeeded')resolve(M('已插入常用语。','定型文を挿入しました。'));else reject(Error(M('请先把光标放在幻灯片的文本框里。','先にスライドのテキストボックスにカーソルを置いてください。')));});}catch(e){reject(e);}});}

// ---- 常用语（三个程序共用）----
var ZH=['此致\n敬礼！','特此通知。','特此证明。','以上报告，请审阅。','如有疑问，请随时与我联系。','感谢您的支持与配合！'],JP=['いつもお世話になっております。','今後ともよろしくお願い申し上げます。','ご確認のほど、よろしくお願いいたします。','以上、よろしくお願いいたします。','取り急ぎご連絡まで。','記','以上'];
function savedPhrases(){try{var a=JSON.parse(root.localStorage.getItem('officeTools.phrases')||'[]');return Array.isArray(a)?a.filter(function(t){return typeof t==='string'&&t;}).slice(0,30):[];}catch(e){return [];}}
var PHRASE_FIELDS=[{key:'phrase',label:'选一句常用语',get options(){var mine=savedPhrases();return mine.map(function(t){return {value:t,label:'★ '+t.replace(/\n/g,' ')};}).concat((ja()?JP:ZH).map(function(t){return {value:t,label:t.replace(/\n/g,' ')};})).concat(mine.length?[{value:'__clear',label:'（清空我保存的常用语）'}]:[]);},get value(){return savedPhrases()[0]||(ja()?JP:ZH)[0];}},{key:'custom',label:'或者自己写一句（会保存下来）',required:false}];
function phrase(host){return async function(api,o){var custom=String(o.custom||'').trim(),t=custom||String(o.phrase||'');if(!custom&&t==='__clear'){root.localStorage.setItem('officeTools.phrases','[]');return M('已清空你保存的常用语（内置的还在）。','保存した定型文を消しました（最初から入っているものは残ります）。');}need(t,M('请选一句常用语，或者自己写一句。','定型文を選ぶか、自分で入力してください。'));need(t.length<=500,M('一句常用语最多 500 字。','定型文は 1 つ 500 文字までです。'));var result=await(host==='excel'?excelPhrase(api,o,t):host==='word'?wordPhrase(api,o,t):pptPhrase(api,o,t));if(custom&&typeof result==='string'){try{root.localStorage.setItem('officeTools.phrases',JSON.stringify([custom].concat(savedPhrases().filter(function(x){return x!==custom;})).slice(0,30)));}catch(e){return result+M('（常用语未能保存，请检查浏览器存储设置。）','（定型文を保存できませんでした。ブラウザーの保存設定を確認してください。）');}}return result;};}

// ---- Office.js 做不到的：给出操作步骤 ----
function guide(zh,jp){return function(){return M(zh,jp);};}
var GUIDES={
  word:{compareDocs:guide('Office 版请用 Word 自带的比较：「审阅 → 比较 → 比较」，原文档选旧版本，修订的文档选新版本。结果在新窗口里，两个原文件都不改。','Word の「校閲 → 比較 → 比較」を使ってください。元の文書に旧版、変更された文書に新版を選ぶと、結果が新しいウィンドウに表示されます。'),
    exportComments:guide('Office 版请用：「文件 → 打印」，在“设置”第一项选“标记列表”，可以打印或存成 PDF。WPS 版能直接生成汇总表。','「ファイル → 印刷」の設定で「変更履歴/コメントの一覧」を選ぶと、印刷や PDF 保存ができます。WPS 版では一覧表を自動で作れます。'),
    captionNumbers:guide('Office 版请用：选中图片或表格 →「引用 → 插入题注」，标签选“图”或“表”。以后增删后按 Ctrl+A、F9 更新编号。','図や表を選んで「参考資料 → 図表番号の挿入」で「図」「表」を選んでください。増減後は Ctrl+A、F9 で番号を更新できます。'),
    fixListNumbering:guide('Office 版请用：把光标放在一级标题上 →「开始 → 多级列表」，选带“标题 1、标题 2”字样的列表样式，所有标题会一起重新编号。','見出し 1 にカーソルを置き、「ホーム → アウトライン」で「見出し 1・見出し 2」と書かれたリストを選ぶと、すべての見出しに番号が振り直されます。')},
  excel:{mergeWorkbooks:guide('Office 版请用 Excel 自带的：「数据 → 获取数据 → 来自文件 → 从文件夹」，选文件夹后点“合并”。WPS 版能直接选文件合并。','Excel の「データ → データの取得 → ファイルから → フォルダーから」でフォルダーを選び、「結合」を押してください。')},
  powerpoint:{slidesToImages:guide('Office 版请用：「文件 → 导出 → 更改文件类型 → PNG 可移植网络图形格式」，选“所有幻灯片”。拼长图请用 WPS 版。','「ファイル → エクスポート → ファイルの種類の変更 → PNG」で「すべてのスライド」を選んでください。縦長画像は WPS 版で作れます。'),
    handoutPdf:guide('Office 版请按 Ctrl+P，把“整页幻灯片”改成“讲义”里的每页 2～9 张，打印机选“Microsoft Print to PDF”可存成 PDF。','Ctrl+P を押し、「フルページサイズのスライド」を「配布資料」（1 ページ 2～9 枚）に変えてください。プリンターで「Microsoft Print to PDF」を選ぶと PDF にできます。'),
    slimDeck:guide('Office 版请用：「视图 → 幻灯片母版」，删掉没用到的版式；再在任意图片上「图片格式 → 压缩图片」，取消“仅应用于此图片”。','「表示 → スライド マスター」で使っていないレイアウトを削除し、画像を選んで「図の形式 → 図の圧縮」で「この画像だけに適用する」を外してください。')}};
var TOOLS={excel:{wanUnit:wanUnit,idCheck:idCheck,attendanceStats:attendanceStats,codesBatch:codesBatch,insertPhrase:phrase('excel')},word:{proofCheck:proofCheck,sendCheck:sendCheck,jaStyle:jaStyle,insertPhrase:phrase('word')},powerpoint:{insertPhrase:phrase('powerpoint')}};
Object.keys(GUIDES).forEach(function(h){Object.keys(GUIDES[h]).forEach(function(id){TOOLS[h][id]=GUIDES[h][id];});});
// 数据脱敏（2026-11-25）：界面在侧栏下面的「数据脱敏」里（mask-panel.js），这里只负责跳过去。
function maskJump(){try{var s=root.document&&root.document.getElementById('maskSection');if(s){s.scrollIntoView({block:'start'});var f=s.querySelector('textarea');if(f)f.focus();}}catch(ignore){}return M('请在侧栏下面的「数据脱敏」里选择范围和方式，先点“找一找”。Office 版会直接改当前文件，请先另存一份副本。','サイドバー下の「データのマスキング」で対象と置き換え方を選び、「探す」を押してください。Office 版は今のファイルを直接変更するので、先にコピーを保存してください。');}
function macroJump(){try{var s=root.document&&root.document.getElementById('macroCheckSection');if(s){s.scrollIntoView({block:'start'});var f=s.querySelector('textarea');if(f)f.focus();}}catch(ignore){}return M('请在侧栏下面的「宏代码检查」里，把 VBA 代码粘贴进去再点“检查”。','サイドバー下の「マクロのチェック」に VBA のコードを貼り付けて「チェック」を押してください。');}
function postalJump(){try{var s=root.document&&root.document.getElementById('postalSection');if(s){s.scrollIntoView({block:'start'});var f=s.querySelector('input');if(f)f.focus();}}catch(ignore){}return M('请在侧栏下面的「郵便番号 ↔ 住所」里输入邮编或住所。','サイドバー下の「郵便番号 ↔ 住所」に郵便番号か住所を入力してください。');}
TOOLS.word.postalOpen=postalJump;TOOLS.excel.postalOpen=postalJump;TOOLS.powerpoint.postalOpen=postalJump;
// 邮编 ↔ 住所：选中一列，右边一列写结果（和 WPS 版规则一样：候选标黄、只到市区町村标浅橙、没找到标浅红）
async function postalFill(api,o){var P=root.OfficePostal;need(P,M('组件没有加载，请重新打开侧栏。','部品が読み込まれていません。サイドバーを開き直してください。'));
  return api.Excel.run(async function(c){var r=c.workbook.getSelectedRange();r.load('values,columnCount,rowCount,columnIndex,rowIndex,worksheet/name');await c.sync();need(r.columnCount===1,M('请只选一列邮编或住所。','郵便番号か住所の列を 1 列だけ選んでください。'));need(r.rowCount<=20000,M('一次最多 2 万行。','一度に 2 万行までです。'));
    var vals=r.values.map(function(row){var x=row[0];if(typeof x==='number')x=String(x).padStart(7,'0');return String(x==null?'':x).trim();}),filled=vals.filter(Boolean);need(filled.length,M('选区是空的。','選択範囲が空です。'));
    var dir=o.dir==='zip2addr'||o.dir==='addr2zip'?o.dir:(filled.filter(P.looksZip).length>=filled.length/2?'zip2addr':'addr2zip');
    var out=r.getOffsetRange(0,1);out.load('values');await c.sync();need(!out.values.some(function(row){return row[0]!==''&&row[0]!=null;}),M('右边一列已经有内容，请先在右边插入一列空白列。','右隣の列にすでに内容があります。先に空の列を挿入してください。'));
    var res=await post(o.fetch,'/api/tools/postal',{action:dir==='zip2addr'?'zip':'address',items:vals}),count={ok:0,multi:0,city:0,none:0,format:0},rows=[],kinds=[];
    res.results.forEach(function(x,i){if(!vals[i]){rows.push(['']);kinds.push('');return;}var t=P.render(x,dir,o.fmt);count[t.kind]++;rows.push([t.t]);kinds.push(t.kind);});
    if(dir==='addr2zip')out.numberFormat=rows.map(function(){return ['@'];});out.values=rows;kinds.forEach(function(k,i){var col=k==='multi'?'#FFFF00':k==='city'?'#FFCC99':k==='none'||k==='format'?'#FFCCCC':'';if(col)out.getCell(i,0).format.fill.color=col;});await c.sync();
    return (dir==='zip2addr'?M('已在右边一列写出住所：','右隣の列に住所を書きました：'):M('已在右边一列写出邮编：','右隣の列に郵便番号を書きました：'))+M('确定 '+count.ok+' 个','確定 '+count.ok+' 件')+(count.multi?M('，有几个候选 '+count.multi+' 个（黄色）','、候補が複数 '+count.multi+' 件（黄色）'):'')+(count.city?M('，只认到市区町村 '+count.city+' 个（浅橙色）','、市区町村まで '+count.city+' 件（薄いオレンジ）'):'')+(count.none+count.format?M('，没找到 '+(count.none+count.format)+' 个（浅红色）','、見つからない '+(count.none+count.format)+' 件（薄い赤）'):'')+M('。数据日期：','。データの日付：')+(res.date||'?');});}
TOOLS.excel.postalFill=postalFill;
TOOLS.word.macroCheck=macroJump;TOOLS.excel.macroCheck=macroJump;TOOLS.powerpoint.macroCheck=macroJump;
TOOLS.word.dataMask=maskJump;TOOLS.excel.dataMask=maskJump;TOOLS.powerpoint.dataMask=guide('PowerPoint 版暂时做不到自动脱敏。请用 WPS 演示打开这份文件，在「办公工具箱 → 数据脱敏」里处理；或者用「开始 → 替换」手动改。','PowerPoint 版では自動マスキングができません。WPS プレゼンテーションで開いて「オフィスツールボックス → データのマスキング」で処理するか、「ホーム → 置換」で手動で直してください。');
var LIST={
  excel:[['postalFill','邮编 ↔ 住所','日本邮编和住所互查，一列批量补全',[{key:'dir',label:'方向',options:[{value:'auto',label:'自动判断'},{value:'zip2addr',label:'邮编 → 住所'},{value:'addr2zip',label:'住所 → 邮编'}],value:'auto'},{key:'fmt',label:'住所写到哪一级',options:[{value:'full',label:'都道府县 + 市区町村 + 町域'},{value:'town',label:'只写町域'}],value:'full'}]],['postalOpen','郵便番号 ↔ 住所','日本邮编查住所、住所查邮编，结果直接插入'],['macroCheck','宏代码检查','粘贴 VBA 代码，找出写错和有风险的地方'],['dataMask','数据脱敏','号码和客户名等换成 * 或代号，可以还原'],['idCheck','号码校验','统一社会信用代码、身份证、发票号、法人番号是否写错',[{key:'type',label:'号码种类',options:[{value:'auto',label:'自动判断'},{value:'uscc',label:'统一社会信用代码'},{value:'id',label:'身份证号'},{value:'invoice',label:'发票号码'},{value:'jpcorp',label:'法人番号 / インボイス登録番号'},{value:'bank',label:'银行卡号'}],value:'auto'}]],
    ['wanUnit','万元显示','金额按万、亿、千、千円显示，数值不变',[{key:'unit',label:'显示成',options:[{value:'wan',label:'万（1 位小数）'},{value:'yi',label:'亿（2 位小数）'},{value:'qian',label:'千（不带单位字）'},{value:'baiwan',label:'百万（不带单位字）'},{value:'senyen',label:'千円'},{value:'general',label:'恢复成普通数字'}],value:'wan'}]],
    ['codesBatch','批量二维码 / 条形码','选中一列网址或编号，每格生成一个二维码或条形码，插在右边',[{key:'kind',label:'生成什么',options:[{value:'qr',label:'二维码（网址、中文都可以）'},{value:'barcode',label:'条形码（Code 128，只能英文和数字）'}],value:'qr'}]],
    ['attendanceStats','考勤统计','按 √、休、假、迟等记号，统计每人各出现几次'],['mergeWorkbooks','合并多个工作簿','（Office 版给出操作步骤）'],['insertPhrase','常用语一键插入','“此致 敬礼”“お世話になっております”等',PHRASE_FIELDS]],
  word:[['postalOpen','郵便番号 ↔ 住所','日本邮编查住所、住所查邮编，结果直接插入'],['macroCheck','宏代码检查','粘贴 VBA 代码，找出写错和有风险的地方'],['dataMask','数据脱敏','号码和客户名等换成 * 或代号，可以还原'],['sendCheck','发送前检查','查批注、修订和作者信息，可一键清理'],['proofCheck','错别字与标点检查','常见错别字、重复字、中英文标点混用、括号不成对，加批注或直接改',[{key:'mode',label:'怎么处理',options:[{value:'comment',label:'加批注提示（不改文字）'},{value:'fix',label:'能确定的直接改，其余加批注'},{value:'clear',label:'清除检查批注'}],value:'comment'}]],
    ['jaStyle','日文文体转换','です・ます体 ⇄ である体，不用 AI',[{key:'to',label:'改成',options:[{value:'plain',label:'である体（です・ます → である）'},{value:'polite',label:'です・ます体（である・だ → です・ます）'}],value:'plain'},{key:'plain',label:'である体的写法',options:[{value:'dearu',label:'〜である（论文、报告）'},{value:'da',label:'〜だ（随笔、内部资料）'}],value:'dearu'}]],
    ['compareDocs','比较两个文档','（Office 版给出操作步骤）'],['exportComments','批注修订汇总导出','（Office 版给出操作步骤）'],['captionNumbers','图表自动编号','（Office 版给出操作步骤）'],['fixListNumbering','修复多级编号','（Office 版给出操作步骤）'],['insertPhrase','常用语一键插入','“此致 敬礼”“お世話になっております”等',PHRASE_FIELDS]],
  powerpoint:[['macroCheck','宏代码检查','粘贴 VBA 代码，找出写错和有风险的地方'],['dataMask','数据脱敏','（Office 版给出操作步骤）'],['insertPhrase','常用语一键插入','“此致 敬礼”“お世話になっております”等',PHRASE_FIELDS],['slidesToImages','每页导出图片 / 长图','（Office 版给出操作步骤）'],['handoutPdf','讲义打印','（Office 版给出操作步骤）'],['slimDeck','PPT 瘦身','（Office 版给出操作步骤）']]};
var oldList=base.list,oldRun=base.run;
base.list=function(host){var have={};var out=oldList(host);out.forEach(function(x){have[x.id]=1;});return out.concat((LIST[host]||[]).filter(function(x){return !have[x[0]];}).map(function(x){return {id:x[0],label:x[1],desc:x[2],form:'',fields:x[3]||null,group:'batch2'};}));};
base.run=function(host,id,api,options){var fn=TOOLS[host]&&TOOLS[host][id];if(fn)return Promise.resolve().then(function(){return fn(api,options||{});});return oldRun(host,id,api,options);};
base.batch2={tools:TOOLS,list:LIST};
})(typeof window!=='undefined'?window:globalThis);
