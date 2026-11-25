// 检查与校对（2026-11-24，第二批新功能）：发送前检查、错别字与标点检查、比较两个文档、批注修订汇总导出、图表自动编号、修复多级编号。
// 只在 WPS 文字里使用。会改文档的都先确认，并合成一步撤销。
(function(root){'use strict';
var T=root.OfficeTools,L=T.lib,P=root.OfficeProof;
function n(v){return Number(v)||0;}
function text(v){return String(v==null?'':v);}
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
// 结果提示直接按界面语言写（数字、文件名夹在中间，整句翻译比逐词翻译通顺）。
function M(zh,jp){return ja()?jp:zh;}
var MARKS=['【办公工具箱】','【オフィスツールボックス】'];function mark(){return MARKS[ja()?1:0];}// 工具加的批注都以这个开头，清除时只删这些
function undoable(a,name,fn){var record;try{try{record=a.UndoRecord;record.StartCustomRecord(name);}catch(ignore){record=null;}return fn();}finally{if(record)try{record.EndCustomRecord();}catch(ignore){}}}
// 在正文里按格式查找全部匹配（隐藏文字、高亮等），返回 [起点, 终点]。
function findAll(d,setup,max){var hits=[],r=d.Content,f=r.Find,last=-1;f.ClearFormatting();setup(f);f.Text='';try{f.Replacement.ClearFormatting();f.Replacement.Text='';}catch(ignore){}f.Forward=true;f.Wrap=0;f.Format=true;f.MatchWildcards=false;
  while(hits.length<(max||500)&&f.Execute()){var s=n(r.Start),e=n(r.End);if(e<=s||s<=last)break;hits.push([s,e]);last=s;r.Collapse(0);}return hits;}
function prop(d,name){try{return text(d.BuiltInDocumentProperties.Item(name).Value).trim();}catch(e){return '';}}
function setProp(d,name,v){try{d.BuiltInDocumentProperties.Item(name).Value=v;return true;}catch(e){return false;}}
function showHidden(d,on){try{var v=d.ActiveWindow.View,old=!!v.ShowHiddenText;v.ShowHiddenText=on;return old;}catch(e){return null;}}

// ---- 发送前检查 ----
function sendCheck(c){var a=c.a,d=c.d,lines=[],comments=0,revs=0,tracking=false,hidden=[],marks=[],people=[];
  try{comments=n(d.Comments.Count);}catch(ignore){}try{revs=n(d.Revisions.Count);}catch(ignore){}try{tracking=!!d.TrackRevisions&&d.TrackRevisions!==0;}catch(ignore){}
  var old=showHidden(d,true);try{hidden=findAll(d,function(f){f.Font.Hidden=true;},200);}catch(ignore){}finally{if(old!==null)showHidden(d,old);}
  try{marks=findAll(d,function(f){f.Highlight=true;},200);}catch(ignore){}
  ['Author','Last Author','Manager','Company'].forEach(function(k){var v=prop(d,k);if(v)people.push({k:k,v:v});});
  var NAMES=ja()?{'Author':'作成者','Last Author':'最終更新者','Manager':'管理者','Company':'会社名'}:{'Author':'作者','Last Author':'最后修改人','Manager':'经理','Company':'单位'};
  if(comments)lines.push(M('批注 '+comments+' 条','コメント '+comments+' 件'));if(revs)lines.push(M('没有接受的修订 '+revs+' 处','未承諾の変更履歴 '+revs+' 件'));if(tracking)lines.push(M('“修订”模式还开着（对方打开会看到修改痕迹）','「変更履歴の記録」がオンのままです（相手に変更の跡が見えます）'));
  if(hidden.length)lines.push(M('隐藏文字 '+hidden.length+' 处（屏幕上看不见，但对方能显示出来）','隠し文字 '+hidden.length+' か所（画面には見えませんが、相手は表示できます）'));if(marks.length)lines.push(M('高亮标记 '+marks.length+' 处','蛍光ペン '+marks.length+' か所'));
  if(people.length)lines.push(M('文件属性里有：','ファイルのプロパティ：')+people.map(function(p){return NAMES[p.k]+M('“'+p.v+'”','「'+p.v+'」');}).join('、'));
  if(!lines.length)return M('检查完毕：没有批注、修订、隐藏文字和高亮，文件属性里也没有作者信息，可以放心发送。','チェック完了：コメント・変更履歴・隠し文字・蛍光ペンはなく、プロパティにも作成者情報はありません。このまま送って大丈夫です。');
  var report=M('发送前检查发现：','送付前チェックの結果：')+'\n· '+lines.join('\n· ');
  var todo=[];if(comments)todo.push(M('删除全部批注','コメントをすべて削除'));if(revs)todo.push(M('接受全部修订','変更履歴をすべて承諾'));if(tracking)todo.push(M('关闭修订模式','変更履歴の記録をオフ'));if(hidden.length)todo.push(M('删除隐藏文字','隠し文字を削除'));if(marks.length)todo.push(M('去掉高亮','蛍光ペンを解除'));if(people.length)todo.push(M('清空作者和单位信息','作成者・会社名を消去'));
  if(!L.ask(report+M('\n\n要一键清理吗？将会：','\n\nまとめて片付けますか？次を行います：')+todo.join('、')+M('。\n批注、修订、隐藏文字和高亮可以用 WPS 撤销恢复；文件属性改了不能撤销。\n选“取消”只看检查结果，不做修改。','。\nコメント・変更履歴・隠し文字・蛍光ペンは WPS の「元に戻す」で戻せます。プロパティの変更は戻せません。\n「キャンセル」なら結果を見るだけで、何も変更しません。')))return report+M('\n\n没有做任何修改。','\n\n何も変更していません。');
  var done=[];undoable(a,'发送前清理',function(){
    if(comments){for(var i=n(d.Comments.Count);i>=1;i--)try{d.Comments.Item(i).Delete();}catch(ignore){}done.push(M('删除批注','コメント削除'));}
    if(revs){try{d.Revisions.AcceptAll();done.push(M('接受修订','変更履歴の承諾'));}catch(e){try{d.AcceptAllRevisions();done.push(M('接受修订','変更履歴の承諾'));}catch(ignore){}}}
    if(tracking)try{d.TrackRevisions=false;done.push(M('关闭修订模式','記録をオフ'));}catch(ignore){}
    if(hidden.length){var o=showHidden(d,true);try{hidden=findAll(d,function(f){f.Font.Hidden=true;},200);for(var j=hidden.length-1;j>=0;j--)try{d.Range(hidden[j][0],hidden[j][1]).Delete();}catch(ignore){}}finally{if(o!==null)showHidden(d,o);}done.push(M('删除隐藏文字','隠し文字削除'));}
    if(marks.length){marks=findAll(d,function(f){f.Highlight=true;},200);marks.forEach(function(m){try{d.Range(m[0],m[1]).HighlightColorIndex=0;}catch(ignore){}});done.push(M('去掉高亮','蛍光ペン解除'));}
  });
  if(people.length){var cleared=people.filter(function(p){return setProp(d,p.k,'');}).length;if(cleared)done.push(M('清空作者信息','作成者情報の消去'));try{d.RemovePersonalInformation=true;}catch(ignore){}}
  var left=[];try{if(n(d.Comments.Count))left.push(M('批注 '+n(d.Comments.Count)+' 条','コメント '+n(d.Comments.Count)+' 件'));if(n(d.Revisions.Count))left.push(M('修订 '+n(d.Revisions.Count)+' 处','変更履歴 '+n(d.Revisions.Count)+' 件'));}catch(ignore){}
  return M('已完成：','完了：')+done.join('、')+M('。','。')+(left.length?M('\n还剩 '+left.join('、')+' 没清理掉，请在“审阅”里手动处理。','\n'+left.join('、')+' が残っています。「校閲」から手動で処理してください。'):'')+M('\n保存后再发送。','\n保存してから送ってください。');}
T.register({id:'sendCheck',label:'发送前检查',desc:'查批注、修订、隐藏文字、高亮和作者信息，可一键清理',icon:'check',group:'proof',undo:false,ownConfirm:['wps'],keywords:'发送前 检查 批注 修订 隐藏文字 作者 属性 清理 痕迹 外发 送付前 チェック コメント 変更履歴 作成者 個人情報',hosts:{wps:sendCheck}});

// ---- 错别字与标点检查 ----
function scopeOf(a,d){var s=a.Selection,t=text(s&&s.Text).replace(/[\r\n\u0007\s]/g,'');if(t.length>1)return {r:s.Range,all:false};return {r:d.Content,all:true};}
function proofCheck(c){var a=c.a,d=c.d,mode=['comment','fix','clear'].indexOf(text((c.options||{}).mode))>=0?c.options.mode:'comment';
  if(mode==='clear'){var gone=0;undoable(a,'清除检查批注',function(){for(var i=n(d.Comments.Count);i>=1;i--){var cm=d.Comments.Item(i);var ct=text(cm.Range.Text);if(MARKS.some(function(m){return ct.indexOf(m)===0;})){try{cm.Delete();gone++;}catch(ignore){}}}});return gone?M('已删除工具加的 '+gone+' 条检查批注，你自己的批注没有动。','ツールが付けたチェック用コメント '+gone+' 件を削除しました。ご自分のコメントはそのままです。'):M('没有找到工具加的检查批注。','ツールが付けたチェック用コメントはありません。');}
  var sc=scopeOf(a,d);if(sc.all&&!L.ask(M('没有选中文字，将检查整篇文档。是否继续？','文字が選択されていないので、文書全体をチェックします。続行しますか？')))return M('已取消。','キャンセルしました。');
  var paras=sc.r.Paragraphs,count=n(paras.Count);if(count>20000)throw Error(M('文档太长（超过 2 万段），请选中一部分再检查。','文書が長すぎます（2 万段落超）。一部を選択してからチェックしてください。'));
  var issues=[];for(var i=1;i<=count;i++){var p=paras.Item(i),r=p.Range,t=text(r.Text),start=n(r.Start);if(!t.replace(/[\r\u0007\s]/g,''))continue;P.check(t.replace(/[\r\u0007]+$/,''),{lang:ja()?'ja':'zh'}).forEach(function(x){x.start+=start;x.end+=start;x.para=i;issues.push(x);});}
  if(!issues.length)return M('检查了 '+count+' 段，没有发现常见错别字和标点问题。（只按规则检查常见错误，不能代替人工校对。）',count+' 段落をチェックし、よくある誤字や句読点の問題は見つかりませんでした。（ルールに基づく簡易チェックです。人の校正の代わりにはなりません。）');
  var MAX=300,fixed=0,noted=0,SUM=P.summary(issues,ja()?'ja':'zh'),MK=mark(),list=issues.slice().sort(function(x,y){return y.start-x.start;});
  if(mode==='fix'&&!L.ask(M('发现 '+issues.length+' 处可能的问题（'+SUM+'）。\n能确定改法的直接改掉，其余的加批注。可以用 WPS 撤销恢复。是否继续？','問題の可能性がある箇所が '+issues.length+' 件あります（'+SUM+'）。\n直し方が決まっているものは直接修正し、残りにはコメントを付けます。WPS の「元に戻す」で戻せます。続行しますか？')))return M('已取消。','キャンセルしました。');
  undoable(a,mode==='fix'?'改正错别字和标点':'检查错别字和标点',function(){list.forEach(function(x){var r=d.Range(x.start,x.end);if(mode==='fix'&&x.fix!=null){try{r.Text=x.fix;fixed++;}catch(ignore){}return;}if(noted>=MAX)return;try{d.Comments.Add(r,MK+x.msg);noted++;}catch(ignore){}});});
  var more=issues.length-fixed-noted>0&&noted>=MAX;
  if(ja())return (mode==='fix'?fixed+' か所を修正しました':'問題の可能性がある箇所が '+issues.length+' 件あります（'+SUM+'）')+(noted?'。'+noted+' か所にコメントを付けたので、1 件ずつ確認してください':'')+(more?'。コメントは最大 '+MAX+' 件なので、直したあともう一度チェックしてください':'')+'。\nコメントは「'+MK+'」で始まります。確認後「チェック用コメントを削除」でまとめて消せます（ご自分のコメントは消えません）。';
  return (mode==='fix'?'已改正 '+fixed+' 处':'发现 '+issues.length+' 处可能的问题（'+SUM+'）')+(noted?'，'+noted+' 处加了批注，请逐条看一下':'')+(more?'；批注最多加 '+MAX+' 条，剩下的请改完后再检查一次':'')+'。\n批注以“'+MK+'”开头；看完后选“清除检查批注”可以一次删掉，你自己的批注不受影响。';}
T.register({id:'proofCheck',label:'错别字与标点检查',desc:'常见错别字、重复字、中英文标点混用、括号不成对，加批注或直接改',icon:'check',group:'proof',undo:false,ownConfirm:['wps'],keywords:'错别字 校对 标点 检查 错字 的的 半角 全角 括号 誤字 脱字 校正 句読点 表記ゆれ',fields:[{key:'mode',label:'怎么处理',options:[{value:'comment',label:'加批注提示（不改文字）'},{value:'fix',label:'能确定的直接改，其余加批注'},{value:'clear',label:'清除检查批注'}],value:'comment'}],hosts:{wps:proofCheck}});

// ---- 日文文体转换（です・ます体 ⇄ である体，2026-11-25）----
// 分词在本机服务里做（kuromoji 离线词典），这里只负责取段落、确认、按位置替换。
async function tool(name,body){var s=await L.localFetch('/api/session',{}).then(function(r){return r.json();});var r=await L.localFetch('/api/tools/'+name,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(body)});var x=await r.json().catch(function(){return {};});if(!r.ok)throw Error(x.error||M('本机服务没有完成，请重试。','ローカルサービスで処理できませんでした。もう一度お試しください。'));return x;}
function clipText(t,k){t=text(t).replace(/[\r\u0007]+/g,'');return t.length>k?t.slice(0,k)+'…':t;}
async function jaStyle(c){var a=c.a,d=c.d,o=c.options||{},to=o.to==='polite'?'polite':'plain',plain=o.plain==='da'?'da':'dearu';
  var sc=scopeOf(a,d);if(sc.all&&!L.ask(M('没有选中文字，将转换整篇文档。是否继续？','文字が選択されていないので、文書全体を変換します。続行しますか？')))return M('已取消。','キャンセルしました。');
  var ps=sc.r.Paragraphs,count=n(ps.Count);if(count>5000)throw Error(M('文档太长（超过 5000 段），请选中一部分再转换。','文書が長すぎます（5000 段落超）。一部を選択してから変換してください。'));
  var list=[];for(var i=1;i<=count;i++){var r=ps.Item(i).Range;list.push({start:n(r.Start),text:text(r.Text).replace(/[\r\u0007]+$/,'')});}
  if(!list.some(function(p){return /[぀-ヿ]/.test(p.text);}))throw Error(M('选中的内容里没有日文。','選択範囲に日本語の文章がありません。'));
  var res=await tool('ja-style',{paragraphs:list.map(function(p){return p.text;}),to:to,plain:plain});
  if(!res.edits&&!res.notes)return to==='plain'?M('没有找到です・ます体的句子，不需要转换。','です・ます体の文は見つかりませんでした。'):M('没有找到である体（だ・である）的句子，不需要转换。','である体（だ・である）の文は見つかりませんでした。');
  var sample=[];res.paragraphs.forEach(function(p,i){if(sample.length>=3||!p.edits.length)return;var t=list[i].text,out=t;p.edits.slice().sort(function(x,y){return y.start-x.start;}).forEach(function(e){out=out.slice(0,e.start)+e.text+out.slice(e.end);});sample.push(clipText(t,40)+'\n  → '+clipText(out,40));});
  var goal=to==='plain'?(plain==='da'?M('だ体','だ体'):M('である体','である体')):M('です・ます体','です・ます体');
  if(!L.ask(M('将把 '+res.edits+' 处改成'+goal+'，例如：\n','次の '+res.edits+' か所を'+goal+'に変えます。例：\n')+sample.join('\n')+(res.notes?M('\n\n另有 '+res.notes+' 处请求句（〜ください 等）改了会不自然，不改，只加批注提示。','\n\n依頼の表現（〜ください など）'+res.notes+' か所は自然に変えられないため、変えずにコメントで知らせます。'):'')+M('\n\n「」里的引用不改。可以用 WPS 撤销恢复。是否继续？','\n\n「」内の引用は変えません。WPS の「元に戻す」で戻せます。続行しますか？')))return M('已取消。','キャンセルしました。');
  var all=[],notes=[];res.paragraphs.forEach(function(p,i){var base=list[i].start;p.edits.forEach(function(e){all.push({s:base+e.start,e:base+e.end,t:e.text,from:e.from});});p.notes.forEach(function(x){notes.push({s:base+x.start,e:base+x.end});});});
  var done=0,skipped=0,mk=mark();
  undoable(a,M('日文文体转换','文体の変換'),function(){
    notes.sort(function(x,y){return y.s-x.s;});all.sort(function(x,y){return y.s-x.s;});
    // 先改文字再加批注会错位：从后往前合并处理
    var steps=all.map(function(x){return {pos:x.s,ed:x};}).concat(notes.map(function(x){return {pos:x.s,note:x};})).sort(function(x,y){return y.pos-x.pos;});
    steps.forEach(function(st){if(st.note){try{d.Comments.Add(d.Range(st.note.s,st.note.e),mk+M('请求句没有自动改。需要的话改成「〜されたい」「〜してほしい」等。','依頼の表現は自動で変えていません。必要なら「〜されたい」「〜してほしい」などに直してください。'));}catch(ignore){}return;}
      var x=st.ed,r=d.Range(x.s,x.e);if(x.from&&text(r.Text)!==x.from){skipped++;return;}try{if(x.t)r.Text=x.t;else r.Delete();done++;}catch(ignore){skipped++;}});});
  return M('已改成'+goal+'：'+done+' 处。',goal+'に変えました：'+done+' か所。')+(skipped?M('\n有 '+skipped+' 处因为文字里有域或隐藏内容，位置对不上，没有改，请手动看一下。','\n'+skipped+' か所はフィールドや隠し文字があり位置が合わないため、変えていません。手動で確認してください。'):'')+(notes.length?M('\n请求句加了 '+notes.length+' 条批注；看完后用“错别字与标点检查 → 清除检查批注”可以一次删掉。','\n依頼の表現に '+notes.length+' 件のコメントを付けました。確認後「誤字・句読点チェック → チェック用コメントを削除」でまとめて消せます。'):'')+M('\n只改句末等处的词尾，建议再通读一遍。','\n文末などの語尾だけを変えています。念のため通して読んでください。');}
T.register({id:'jaStyle',label:'日文文体转换',desc:'です・ます体 ⇄ である体，不用 AI，只改句末，引用不动',icon:'type',group:'proof',undo:false,ownConfirm:['wps'],keywords:'文体 です ます である だ 敬体 常体 変換 日文 日语 口調 論文 レポート 文末',fields:[{key:'to',label:'改成',options:[{value:'plain',label:'である体（です・ます → である）'},{value:'polite',label:'です・ます体（である・だ → です・ます）'}],value:'plain'},{key:'plain',label:'である体的写法',options:[{value:'dearu',label:'〜である（论文、报告）'},{value:'da',label:'〜だ（随笔、内部资料）'}],value:'dearu'}],hosts:{wps:jaStyle}});

// ---- 比较两个文档 ----
function pickOne(a,title){var dlg=a.FileDialog(3);dlg.Title=root.OfficeI18n?root.OfficeI18n.t(title):title;dlg.AllowMultiSelect=false;try{dlg.Filters.Clear();dlg.Filters.Add(M('文档','文書'),'*.docx;*.doc;*.wps;*.rtf');}catch(ignore){}if(dlg.Show()!==-1)return '';return text(dlg.SelectedItems.Item(1));}
function compareDocs(c){var a=c.a,d=c.d,file=pickOne(a,'选择要和当前文档比较的另一个版本（修改后的）');if(!file)return M('已取消。','キャンセルしました。');
  if(file.toLowerCase()===text(d.FullName).toLowerCase())throw Error(M('选的是当前文档本身，请选另一个版本。','今開いている文書が選ばれています。別の版を選んでください。'));
  var how='';try{d.Compare(file,'办公工具箱',2,true,true,false,false,false);how='compare';}catch(e1){
    var other=null;try{other=a.Documents.Open(file,false,true,false);a.CompareDocuments(d,other,2,1,true,true,true,true,true,true,true,true,true,true,'办公工具箱',true);how='compareDocuments';}catch(e2){how='';}finally{try{if(other)other.Close(0);}catch(ignore){}}}
  var fname=file.replace(/^.*[\\\/]/,'');if(!how)throw Error(M('这台 WPS 不支持自动比较。请用 WPS 顶部的“审阅 → 比较”，原文档选当前文档，修订的文档选“'+fname+'”。','この WPS では自動比較が使えません。上部の「校閲 → 比較」で、元の文書に今の文書、変更された文書に「'+fname+'」を選んでください。'));
  return M('已在新窗口生成比较结果：增加的文字带下划线，删除的带删除线，右边列出每一处改动。两个原文件都没有改动；比较结果需要的话请另存。','比較結果を新しいウィンドウに作成しました。追加は下線、削除は取り消し線で表示され、右側に変更箇所の一覧があります。元の 2 つのファイルは変更していません。結果が必要なら名前を付けて保存してください。');}
T.register({id:'compareDocs',label:'比较两个文档',desc:'选另一个版本，生成标出所有改动的新文档',icon:'match',group:'proof',undo:false,ownConfirm:['wps'],keywords:'比较 对比 两个版本 改了哪里 差异 合同 比較 差分 違い 新旧',hosts:{wps:compareDocs}});

// ---- 批注修订汇总导出 ----
function when(v){if(v==null||v==='')return '';var dt=v instanceof Date?v:typeof v==='number'?new Date(Math.round((v-25569)*86400000)):new Date(String(v));if(isNaN(dt))return text(v);var p=function(x){return String(x).padStart(2,'0');};return dt.getFullYear()+'-'+p(dt.getMonth()+1)+'-'+p(dt.getDate())+' '+p(dt.getHours())+':'+p(dt.getMinutes());}
function page(r){try{return n(r.Information(3))||'';}catch(e){return '';}}
function clip(s,k){s=text(s).replace(/[\r\u0007]+/g,' ').replace(/\s+/g,' ').trim();return s.length>k?s.slice(0,k)+'…':s;}
var REV={1:'插入',2:'删除',3:'属性',4:'段落编号',5:'域显示',6:'对照',7:'冲突',8:'样式',9:'替换',10:'段落格式',11:'表格属性',12:'节属性',13:'样式定义',14:'移出',15:'移入',16:'单元格插入',17:'单元格删除',18:'单元格合并',19:'单元格拆分'};
function exportComments(c){var a=c.a,d=c.d,J=ja(),rows=[],k=0,nc=0,nr=0;
  for(var i=1;i<=n(d.Comments.Count)&&rows.length<1000;i++){var cm=d.Comments.Item(i),scope=null;try{scope=cm.Scope;}catch(ignore){}rows.push([String(++k),J?'コメント':'批注',String(scope?page(scope):''),text(cm.Author),when(cm.Date),clip(scope&&scope.Text,80),clip(cm.Range.Text,300)]);nc++;}
  var rv=0;try{rv=n(d.Revisions.Count);}catch(ignore){}
  for(var j=1;j<=rv&&rows.length<1000;j++){var r=d.Revisions.Item(j),type=n(r.Type);rows.push([String(++k),(J?{1:'挿入',2:'削除'}[type]:REV[type])||(J?'書式':'格式'),String(page(r.Range)),text(r.Author),when(r.Date),'',clip(r.Range.Text,300)]);nr++;}
  if(!rows.length)return M('这份文档里没有批注，也没有修订。','この文書にはコメントも変更履歴もありません。');
  var nd=a.Documents.Add(),name=text(d.Name),title=(J?'コメント・変更履歴一覧：':'批注和修订汇总：')+name;
  nd.Content.Text=title+'\r'+(J?'作成：':'导出时间：')+when(new Date())+(J?'　コメント '+nc+' 件、変更 '+nr+' 件':'　批注 '+nc+' 条、修订 '+nr+' 处')+'\r';
  try{var hp=nd.Paragraphs.Item(1).Range;hp.Font.Size=14;hp.Font.Bold=true;}catch(ignore){}
  var at=nd.Content;at.Collapse(0);var head=J?['No.','種類','ページ','作成者','日時','対象の文章','内容']:['序号','类型','页码','作者','时间','原文','批注 / 修改内容'];
  var t=L.addWordTable(nd,at,[head].concat(rows));try{t.Rows.Item(1).Range.Font.Bold=true;t.Rows.Item(1).HeadingFormat=-1;t.AutoFitBehavior(2);}catch(ignore){}
  return M('已在新文档里列出 '+nc+' 条批注、'+nr+' 处修订（共 '+rows.length+' 行）。'+(rows.length>=1000?'最多列 1000 行。':'')+'原文档没有改动；新文档需要的话请保存。','新しい文書にコメント '+nc+' 件、変更履歴 '+nr+' 件（計 '+rows.length+' 行）を一覧にしました。'+(rows.length>=1000?'一覧は最大 1000 行です。':'')+'元の文書は変更していません。必要なら新しい文書を保存してください。');}
T.register({id:'exportComments',label:'批注修订汇总导出',desc:'把全部批注和修订列成一张表，放在新文档里',icon:'list',group:'proof',undo:false,ownConfirm:['wps'],keywords:'批注 修订 汇总 导出 列表 意见 审阅 コメント 一覧 変更履歴 エクスポート',hosts:{wps:exportComments}});

// ---- 图表自动编号 ----
var CAP=/^\s*(图|表|図)\s*(\d+)(?=\s|[.．、:：\-—–]|[^\d]|$)/;
function capLabels(d){var sample='';try{sample=text(d.Content.Text).slice(0,5000);}catch(ignore){}var jp=/[぀-ヿ]/.test(sample);return {pic:jp?'図':'图',table:'表'};}
function addField(d,pos,label){var r=d.Range(pos,pos);return d.Fields.Add(r,-1,'SEQ '+label+' \\* ARABIC',false);}
function captionStyle(d,pos){try{var p=d.Range(pos,pos).Paragraphs.Item(1);try{p.Style=-35;}catch(ignore){}p.Alignment=1;}catch(ignore){}}
function captionNumbers(c){var a=c.a,d=c.d,o=c.options||{},what=['both','pictures','tables'].indexOf(o.what)>=0?o.what:'both',lab=capLabels(d),jobs=[],renum=[],paras=d.Paragraphs,count=n(paras.Count);
  if(count>50000)throw Error(M('文档太长，暂不支持。','文書が長すぎるため対応していません。'));
  for(var i=1;i<=count;i++){var p=paras.Item(i),t=text(p.Range.Text),m=CAP.exec(t);if(!m)continue;var hasField=false;try{hasField=n(p.Range.Fields.Count)>0;}catch(ignore){}if(hasField)continue;if(t.replace(/[\r\u0007]/g,'').length>120)continue;
    var kind=m[1]==='表'?'table':'pic';if(what==='pictures'&&kind==='table'||what==='tables'&&kind==='pic')continue;var off=m[0].length-m[2].length,s=n(p.Range.Start)+off;renum.push({s:s,e:s+m[2].length,label:m[1]});}
  function isCaption(t,kind){var m=CAP.exec(text(t));return !!m&&(kind==='table'?m[1]==='表':m[1]!=='表');}
  if(what!=='pictures'){for(var k=1;k<=n(d.Tables.Count);k++){var tb=d.Tables.Item(k);try{if(n(tb.NestingLevel)>1)continue;}catch(ignore){}var ts=n(tb.Range.Start);if(ts<=0)continue;var prev=d.Range(ts-1,ts-1).Paragraphs.Item(1);if(isCaption(prev.Range.Text,'table'))continue;jobs.push({pos:ts,kind:'table',para:prev});}}
  if(what!=='tables'){var shapes=d.InlineShapes;for(var q=1;q<=n(shapes.Count);q++){var sh=shapes.Item(q),ty=n(sh.Type);if(ty!==3&&ty!==4)continue;var sr=sh.Range;try{if(sr.Information(12))continue;}catch(ignore){}var ip=sr.Paragraphs.Item(1),end=n(ip.Range.End),next=d.Range(end,end).Paragraphs.Item(1);if(n(next.Range.Start)===end&&isCaption(next.Range.Text,'pic'))continue;if(jobs.some(function(j){return j.kind==='pic'&&j.pos===end;}))continue;jobs.push({pos:end,kind:'pic',para:ip});}}
  if(!jobs.length&&!renum.length)return M('没有需要编号的图片或表格（已有题注的不会重复加）。','番号を付ける図や表はありません（キャプションがあるものには付けません）。');
  var np=jobs.filter(function(j){return j.kind==='pic';}).length,nt=jobs.filter(function(j){return j.kind==='table';}).length;
  if(!L.ask(M('将给 '+np+' 张图片加图题（在图片下面）、'+nt+' 个表格加表题（在表格上面）'+(renum.length?'，并把 '+renum.length+' 处手写的编号改成自动编号':'')+'。\n编号会按顺序自动排，以后增删图表也不会乱。可以用 WPS 撤销恢复。是否继续？','図 '+np+' 個（図の下）と表 '+nt+' 個（表の上）にキャプションを付けます'+(renum.length?'。手入力の番号 '+renum.length+' か所も自動番号に変えます':'')+'。\n番号は順番に自動で振られ、図表を増減しても崩れません。WPS の「元に戻す」で戻せます。続行しますか？')))return M('已取消。','キャンセルしました。');
  var added={pic:0,table:0},changed=0;
  undoable(a,'图表自动编号',function(){
    var all=jobs.map(function(j){return {pos:j.pos,job:j};}).concat(renum.map(function(r){return {pos:r.s,ren:r};})).sort(function(x,y){return y.pos-x.pos;});
    all.forEach(function(x){if(x.ren){try{d.Fields.Add(d.Range(x.ren.s,x.ren.e),-1,'SEQ '+x.ren.label+' \\* ARABIC',false);changed++;}catch(ignore){}return;}
      var j=x.job,label=j.kind==='table'?lab.table:lab.pic;try{
        var para=j.para,blank=!text(para.Range.Text).replace(/[\r\u0007\s]/g,'')&&j.kind==='table',at;
          if(blank)at=n(para.Range.Start);else{at=n(para.Range.End);d.Range(at-1,at-1).InsertBefore('\r');}
        // 新段落可能继承了图片段落的格式，先写文字再加编号
        var r=d.Range(at,at);r.InsertAfter(label+' ');addField(d,at+label.length+1,label);captionStyle(d,at);added[j.kind]++;}catch(ignore){}});
    try{d.Fields.Update();}catch(ignore){}});
  return M('已加上 '+added.pic+' 个图题、'+added.table+' 个表题'+(changed?'，'+changed+' 处手写编号改成了自动编号':'')+'。\n请在编号后面补上图表的名称。以后增删图表后，按 Ctrl+A 全选再按 F9，编号会重新排好。','図のキャプション '+added.pic+' 個、表のキャプション '+added.table+' 個を付けました'+(changed?'。手入力の番号 '+changed+' か所を自動番号にしました':'')+'。\n番号の後ろに図表の名前を書き足してください。図表を増減したら Ctrl+A で全選択して F9 を押すと番号が振り直されます。');}
T.register({id:'captionNumbers',label:'图表自动编号',desc:'给图片、表格加“图 1”“表 1”题注，手写编号改成自动编号',icon:'list',group:'proof',undo:false,ownConfirm:['wps'],keywords:'题注 图表编号 图1 表1 自动编号 插图 表格编号 図表番号 キャプション 図1 表1',fields:[{key:'what',label:'给哪些加编号',options:[{value:'both',label:'图片和表格'},{value:'pictures',label:'只给图片'},{value:'tables',label:'只给表格'}],value:'both'}],hosts:{wps:captionNumbers}});

// ---- 修复多级编号 ----
var MANUAL=/^\s*(?:第[一二三四五六七八九十百零〇\d]+[章节条部篇]|[一二三四五六七八九十]+[、．.]|[（(][一二三四五六七八九十\d]+[）)]|\d+(?:[.．]\d+){0,3}(?:[.．、]|[\s　]+)|[①-⑳])[\s　]*/;
var SCHEMES={arabic:[['%1',0],['%1.%2',0],['%1.%2.%3',0]],chinese:[['%1、',37],['（%2）',37],['%3.',0]],chapter:[['第%1章',0],['%1.%2',0],['%1.%2.%3',0]]};
function fixListNumbering(c){var a=c.a,d=c.d,o=c.options||{},scheme=SCHEMES[o.scheme]?o.scheme:'arabic',strip=o.strip!=='no',paras=d.Paragraphs,count=n(paras.Count),heads=[],other=0;
  var styleNames=[1,2,3].map(function(l){try{return text(d.Styles.Item(-1-l).NameLocal);}catch(e){return '';}});
  for(var i=1;i<=count&&i<=50000;i++){var p=paras.Item(i),lvl=0;try{lvl=n(p.OutlineLevel);}catch(ignore){}if(lvl<1||lvl>3)continue;var sn='';try{sn=text(p.Style&&p.Style.NameLocal!=null?p.Style.NameLocal:p.Style);}catch(ignore){}
    if(styleNames.indexOf(sn)<0){other++;continue;}heads.push({p:p,lvl:lvl});}
  if(!heads.length)throw Error(M('没有找到用“标题 1 / 标题 2 / 标题 3”样式的标题。请先用“一级标题”“二级标题”按钮把标题设好，再修复编号。','「見出し 1～3」スタイルの見出しが見つかりません。先に「見出し 1」「見出し 2」ボタンで見出しを設定してから、もう一度お試しください。'));
  var manual=strip?heads.filter(function(h){var lt=0;try{lt=n(h.p.Range.ListFormat.ListType);}catch(ignore){}return !lt&&MANUAL.test(text(h.p.Range.Text));}):[];
  var sample={arabic:'1 / 1.1 / 1.1.1',chinese:'一、/（一）/ 1.',chapter:'第1章 / 1.1 / 1.1.1'}[scheme],cnt=[0,0,0];heads.forEach(function(h){cnt[h.lvl-1]++;});
  if(!L.ask(M('将给 '+cnt[0]+' 个一级、'+cnt[1]+' 个二级、'+cnt[2]+' 个三级标题重新自动编号（'+sample+'），下级跟着上级重新从 1 开始。'+(manual.length?'\n标题开头手打的 '+manual.length+' 处编号会去掉，避免出现两个编号。':'')+'\n可以用 WPS 撤销恢复。是否继续？','見出し 1 を '+cnt[0]+' 個、見出し 2 を '+cnt[1]+' 個、見出し 3 を '+cnt[2]+' 個、自動番号で振り直します（'+sample+'）。下の階層は上が変わると 1 から数え直します。'+(manual.length?'\n見出しの先頭に手入力した番号 '+manual.length+' か所は、二重にならないよう削除します。':'')+'\nWPS の「元に戻す」で戻せます。続行しますか？')))return M('已取消。','キャンセルしました。');
  var removed=0;undoable(a,'修复多级编号',function(){
    manual.slice().reverse().forEach(function(h){var r=h.p.Range,m=MANUAL.exec(text(r.Text));if(!m||!m[0])return;try{d.Range(n(r.Start),n(r.Start)+m[0].length).Delete();removed++;}catch(ignore){}});
    var lt=d.ListTemplates.Add(true),S=SCHEMES[scheme];
    for(var k=0;k<3;k++){var lv=lt.ListLevels.Item(k+1);try{lv.NumberFormat=S[k][0];}catch(ignore){}try{lv.NumberStyle=S[k][1];}catch(ignore){}try{lv.TrailingCharacter=1;}catch(ignore){}try{lv.StartAt=1;}catch(ignore){}try{if(k)lv.ResetOnHigher=k;}catch(ignore){}try{lv.NumberPosition=0;lv.TextPosition=0;lv.TabPosition=0;}catch(ignore){}try{lv.LinkedStyle=styleNames[k];}catch(ignore){}}
    for(var k2=0;k2<3;k2++){try{d.Styles.Item(-2-k2).LinkToListTemplate(lt,k2+1);}catch(ignore){}}
    heads.forEach(function(h){try{if(!n(h.p.Range.ListFormat.ListType))h.p.Range.ListFormat.ApplyListTemplateWithLevel(lt,true,1,2,h.lvl);}catch(ignore){}});});
  var missing=heads.filter(function(h){try{return !n(h.p.Range.ListFormat.ListType);}catch(e){return false;}}).length;
  if(ja())return '番号を振り直しました：見出し 1 が '+cnt[0]+' 個、見出し 2 が '+cnt[1]+' 個、見出し 3 が '+cnt[2]+' 個（'+sample+'）'+(removed?'。手入力の番号 '+removed+' か所を削除しました':'')+'。'+(other?'\n「見出し 1～3」スタイルでない見出しが '+other+' 個あり、変更していません。「見出し 1」「見出し 2」ボタンで設定してからもう一度実行してください。':'')+(missing?'\n番号が付かなかった見出しが '+missing+' 個あります。確認してください。':'');
  return '已重新编号：一级 '+cnt[0]+' 个、二级 '+cnt[1]+' 个、三级 '+cnt[2]+' 个（'+sample+'）'+(removed?'，去掉了 '+removed+' 处手打的编号':'')+'。'+(other?'\n另有 '+other+' 个标题没有用“标题 1～3”样式，没有改，可以先用“一级标题”“二级标题”按钮设好再点一次。':'')+(missing?'\n有 '+missing+' 个标题没加上编号，请检查。':'');}
T.register({id:'fixListNumbering',label:'修复多级编号',desc:'标题编号乱了、不连续、手打的：按标题级别重新自动编号',icon:'list',group:'proof',undo:false,ownConfirm:['wps'],keywords:'多级编号 编号乱 编号不连续 标题编号 1.1 自动编号 章节编号 段落番号 見出し番号 アウトライン 番号がずれる',fields:[{key:'scheme',label:'编号样式',options:[{value:'arabic',label:'1 / 1.1 / 1.1.1'},{value:'chinese',label:'一、/（一）/ 1.'},{value:'chapter',label:'第1章 / 1.1 / 1.1.1'}],value:'arabic'},{key:'strip',label:'标题开头手打的编号',options:[{value:'yes',label:'去掉（推荐）'},{value:'no',label:'保留'}],value:'yes'}],hosts:{wps:fixListNumbering}});
})(window);
