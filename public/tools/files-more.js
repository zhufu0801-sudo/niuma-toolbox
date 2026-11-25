// 第二批新功能（2026-11-24）：批量重命名文件、常用语一键插入。文字、表格、演示里都能用。
(function(root){'use strict';
var T=root.OfficeTools,L=T.lib;
function n(v){return Number(v)||0;}
function text(v){return String(v==null?'':v);}
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
function tt(s){return root.OfficeI18n?root.OfficeI18n.t(s):s;}
var CANCEL=function(){return M('已取消。','キャンセルしました。');};
async function tool(name,body){var s=await L.localFetch('/api/session',{}).then(function(r){return r.json();});var r=await L.localFetch('/api/tools/'+name,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(body)});var d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(d.error||M('本机服务没有完成，请重试。','ローカルサービスで処理できませんでした。もう一度お試しください。'));return d;}
function base(p){return text(p).replace(/^.*[\\\/]/,'');}

// ---- 批量重命名 ----
async function batchRename(c){var a=c.a,o=c.options||{},mode=['prefix','suffix','replace','number','undo'].indexOf(o.mode)>=0?o.mode:'prefix';
  if(mode==='undo'){if(!L.ask(M('将把最近一次批量重命名的文件改回原来的名字。是否继续？','直前の一括名前変更を元の名前に戻します。続行しますか？')))return CANCEL();var u=await tool('rename',{action:'undo'});return M('已改回 '+u.back.length+' 个文件。',u.back.length+' 個のファイルを元の名前に戻しました。')+(u.failed.length?M('\n没改回的：','\n戻せなかったもの：')+u.failed.slice(0,5).map(function(x){return base(x.from)+'（'+tt(x.error)+'）';}).join(M('；','、')):'');}
  var rule={mode:mode,text:text(o.text),find:text(o.find),start:text(o.start)};if(mode==='replace'&&!rule.find)throw Error(M('请填写“要替换的文字”。','「置き換える文字」を入力してください。'));if((mode==='prefix'||mode==='suffix')&&!rule.text)throw Error(M('请填写要加的文字。','付け足す文字を入力してください。'));
  var dlg=a.FileDialog(3);dlg.Title=M('选择要改名的文件（可多选）','名前を変えるファイルを選択（複数選択可）');dlg.AllowMultiSelect=true;try{dlg.Filters.Clear();dlg.Filters.Add(M('所有文件','すべてのファイル'),'*.*');}catch(ignore){}if(dlg.Show()!==-1)return CANCEL();
  var paths=[];for(var i=1;i<=n(dlg.SelectedItems.Count);i++)paths.push(text(dlg.SelectedItems.Item(i)));paths.sort(function(x,y){return x.localeCompare(y,'zh-CN',{numeric:true});});
  var pre=await tool('rename',{action:'preview',paths:paths,rule:rule}),ok=pre.plan.filter(function(x){return !x.skip;}),skip=pre.plan.filter(function(x){return x.skip;});
  if(!ok.length)throw Error(M('没有可以改名的文件：','名前を変えられるファイルがありません：')+skip.slice(0,3).map(function(x){return base(x.from)+'（'+tt(x.skip)+'）';}).join(M('；','、')));
  var sample=ok.slice(0,8).map(function(x){return base(x.from)+' → '+base(x.to);}).join('\n');
  if(!L.ask(M('将给 '+ok.length+' 个文件改名，例如：\n',ok.length+' 個のファイルの名前を変えます。例：\n')+sample+(ok.length>8?'\n……':'')+(skip.length?M('\n另有 '+skip.length+' 个不改（'+skip[0].skip+(skip.length>1?' 等':'')+'）','\nほかの '+skip.length+' 個は変えません（'+tt(skip[0].skip)+(skip.length>1?' など':'')+'）'):'')+M('\n\n改完会留一份记录，可以用“改回上一次”恢复。是否继续？','\n\n記録を残すので、「直前の変更を元に戻す」で戻せます。続行しますか？')))return CANCEL();
  var res=await tool('rename',{action:'apply',paths:paths,rule:rule});
  return M('已改名 '+res.done.length+' 个文件。',res.done.length+' 個のファイルの名前を変えました。')+(res.failed.length?M('\n没改成的：','\n変えられなかったもの：')+res.failed.slice(0,5).map(function(x){return base(x.from)+'（'+tt(x.error)+'）';}).join(M('；','、')):'')+(res.log?M('\n改错了可以在这里选“改回上一次”。','\n間違えたときは「直前の変更を元に戻す」を選んでください。'):'');}
var RENAME_FIELDS=[{key:'mode',label:'怎么改名',options:[{value:'prefix',label:'前面加文字'},{value:'suffix',label:'后面加文字'},{value:'replace',label:'把一段文字换成另一段'},{value:'number',label:'按顺序编号（01、02…）'},{value:'undo',label:'改回上一次'}],value:'prefix',required:true},
  {key:'text',label:'要加的文字 / 换成的文字 / 编号格式',placeholder:'例如：2026-；编号时写“资料-#”，# 是编号的位置',required:false},
  {key:'find',label:'要替换的文字（只在“换文字”时填）',placeholder:'例如：初稿',required:false},
  {key:'start',label:'编号从几开始（不填从 1）',placeholder:'1',required:false}];
T.register({id:'batchRename',label:'批量重命名',desc:'一次给很多文件加前缀后缀、换文字或按顺序编号，可以改回',icon:'replace',group:'file',undo:false,ownConfirm:['wps','et','wpp'],keywords:'批量重命名 改名 文件名 编号 前缀 后缀 一括 リネーム ファイル名 変更 連番',fields:RENAME_FIELDS,hosts:{wps:batchRename,et:batchRename,wpp:batchRename}});

// ---- 常用语 ----
var ZH=['此致\n敬礼！','特此通知。','特此证明。','以上报告，请审阅。','如有疑问，请随时与我联系。','感谢您的支持与配合！','请于    年  月  日前回复。','附件：','（此页无正文）','（以下无正文）'];
var JP=['いつもお世話になっております。','今後ともよろしくお願い申し上げます。','ご確認のほど、よろしくお願いいたします。','以上、よろしくお願いいたします。','ご不明な点がございましたら、お気軽にお問い合わせください。','取り急ぎご連絡まで。','拝啓　時下ますますご清栄のこととお慶び申し上げます。','敬具','記','以上'];
var KEY='officeTools.phrases';
function storage(){try{var p=T.app().PluginStorage;if(p&&typeof p.getItem==='function')return p;}catch(ignore){}try{return root.localStorage;}catch(ignore){}return null;}
function saved(){try{var v=JSON.parse(storage().getItem(KEY)||'[]');return Array.isArray(v)?v.filter(function(x){return typeof x==='string'&&x;}).slice(0,30):[];}catch(e){return [];}}
function save(list){try{storage().setItem(KEY,JSON.stringify(list.slice(0,30)));}catch(ignore){}}
function short(t){t=t.replace(/\n/g,' ⏎ ');return t.length>40?t.slice(0,40)+'…':t;}
// 选项每次打开表单时重新生成：先列自己保存的，再列内置的（按界面语言）。
var PHRASE_FIELD={key:'phrase',label:'选一句常用语',required:true,get options(){var mine=saved().map(function(t){return {value:t,label:'★ '+short(t)};}),built=(ja()?JP:ZH).map(function(t){return {value:t,label:short(t)};});return mine.concat(built).concat(mine.length?[{value:'__clear',label:'（清空我保存的常用语）'}]:[]);},get value(){var m=saved();return m[0]||(ja()?JP[0]:ZH[0]);}};
function pickPhrase(o){var custom=text(o.custom).replace(/\r\n?/g,'\n').trim();if(custom){if(custom.length>500)throw Error(M('一句常用语最多 500 字。','定型文は 1 つ 500 文字までです。'));var list=saved().filter(function(x){return x!==custom;});list.unshift(custom);save(list);return {text:custom,saved:true};}
  var p=text(o.phrase);if(p==='__clear')return {clear:true};if(!p)throw Error(M('请选一句常用语，或者自己写一句。','定型文を選ぶか、自分で入力してください。'));return {text:p};}
function insertPhrase(c){var a=c.a,h=c.h,x=pickPhrase(c.options||{});if(x.clear){save([]);return M('已清空你保存的常用语（内置的还在）。','保存した定型文を消しました（最初から入っているものは残ります）。');}
  var t=x.text,note=x.saved?M('（已保存，下次在列表最上面）','（保存しました。次回から一覧のいちばん上に出ます）'):'',DONE=M('已插入常用语','定型文を挿入しました');
  if(h==='wps'){var r=a.Selection.Range;if(n(r.Start)!==n(r.End)){r=a.Selection.Range.Duplicate;r.Collapse(0);}var rec;try{try{rec=a.UndoRecord;rec.StartCustomRecord('插入常用语');}catch(ignore){rec=null;}r.Text=t.replace(/\n/g,'\r');try{r.Select();a.Selection.Collapse(0);}catch(ignore){}}finally{if(rec)try{rec.EndCustomRecord();}catch(ignore){}}return DONE+note+M('。','');}
  if(h==='et'){var cell=L.cells(a).Cells.Item(1,1);if(text(cell.Value2)!==''&&!L.ask(M('选中的格子里已经有内容，要换成这句常用语吗？','選んだセルにはすでに内容があります。この定型文に置き換えますか？')))return CANCEL();cell.Value2=/^[=+@-]/.test(t)?"'"+t:t;if(/\n/.test(t))try{cell.WrapText=true;}catch(ignore){}return M('已写入常用语','定型文を入力しました')+note+M('。','');}
  var sel=a.ActiveWindow.Selection;if(n(sel.Type)!==3)throw Error(M('请先把光标放在幻灯片的文本框里。','先にスライドのテキストボックスにカーソルを置いてください。'));var tr=sel.TextRange,ins=t.replace(/\n/g,'\r');if(tr.InsertAfter)tr.InsertAfter(ins);else tr.Text=text(tr.Text)+ins;return DONE+note+M('。','');}
T.register({id:'insertPhrase',label:'常用语一键插入',desc:'“此致 敬礼”“お世話になっております”等，也可以存自己常用的话',icon:'insert',group:'stamp',undo:false,ownConfirm:['wps','et','wpp'],keywords:'常用语 套话 结尾语 问候语 此致敬礼 特此通知 定型文 挨拶 お世話になっております よろしくお願いいたします 敬具',fields:[PHRASE_FIELD,{key:'custom',label:'或者自己写一句（会保存下来）',placeholder:'例如：请各部门于周五前提交。',required:false}],hosts:{wps:insertPhrase,et:insertPhrase,wpp:insertPhrase}});
})(window);
