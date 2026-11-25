(function(root){'use strict';
// 按名单批量生成（侧栏“文件与内容”页，WPS 文字里显示）：当前文档当模板，选一个 Excel 名单，每行生成一份。
var T=root.OfficeTools,box=null,listPath=null,status=null,info=null,nameSel=null,busy=false;
function t(s){return root.OfficeI18n?root.OfficeI18n.t(s):s;}
function el(tag,cls,text){var e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;}
// 状态行在这里翻译好再显示（里面有列名、文件位置等用户内容，不能交给整页翻译逐词替换）。
function say(m,bad,raw){status.textContent=raw?m:t(m);status.className='mm-status'+(bad?' bad':'');}
async function api(body){var L=T.lib,s=await L.localFetch('/api/session',{}).then(function(r){if(!r.ok)throw Error('本地服务未连接。');return r.json();});var r=await L.localFetch('/api/tools/mail-merge',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(body)});var d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(d.error||'没有完成，请重试。');return d;}
// 模板 = 当前文档。有改动没保存时先问一句再保存；从没保存过的新文档要先另存。
function templatePath(){var a=T.app(),d=a&&a.ActiveDocument;if(!d)throw Error('请先打开当作模板的 WPS 文字文档。');var full=String(d.FullName||'');
  if(!/[\\\/]/.test(full))throw Error('请先把这个文档保存成 .docx（文件 → 另存为），再来生成。');if(!/\.docx$/i.test(full))throw Error('模板要是 .docx 格式，请另存为 Word 文档（.docx）后再试。');
  if(d.Saved===false||d.Saved===0){if(!T.lib.ask('模板有改动还没保存。先保存再继续吗？'))throw Error('已取消。');d.Save();}return full;}
function pickList(){try{var a=T.app(),dlg=a.FileDialog(3);dlg.Title=t('选择名单（Excel 表格）');dlg.AllowMultiSelect=false;try{dlg.Filters.Clear();dlg.Filters.Add(t('Excel 表格'),'*.xlsx');}catch(ignore){}if(dlg.Show()!==-1)return null;return String(dlg.SelectedItems.Item(1));}catch(e){return undefined;}}
function task(fn){if(busy)return;busy=true;box.querySelectorAll('button').forEach(function(b){b.disabled=true;});Promise.resolve().then(fn).catch(function(e){say(e.message==='已取消。'?t(e.message):t('未完成：')+t(e.message),true,true);}).then(function(){busy=false;box.querySelectorAll('button').forEach(function(b){b.disabled=false;});});}
function fillColumns(cols){var keep=nameSel.value;nameSel.textContent='';var o=el('option','','（按顺序编号）');o.value='';nameSel.append(o);cols.forEach(function(c){var op=document.createElement('option');op.value=c;op.textContent=c;op.setAttribute('translate','no');nameSel.append(op);});if(cols.indexOf(keep)>=0)nameSel.value=keep;else{var guess=cols.find(function(c){return /姓名|名前|氏名|名称|会社|公司|客户/.test(c);});if(guess)nameSel.value=guess;}}
function check(){return task(async function(){var tpl=templatePath(),list=listPath.value.trim();if(!list)throw Error('请先选择名单文件。');say('正在读取名单和模板…');var d=await api({template:tpl,list:list,preview:true});fillColumns(d.columns);
  var lines=[t('名单共 '+d.count+' 行。'),t('模板里用到的列：')+(d.used.length?d.used.join('、'):t('（没有）'))];if(d.unused.length)lines.push(t('名单里没用到的列：')+d.unused.join('、'));if(d.unknown.length)lines.push(t('模板里有、名单里没有（不会替换）：')+d.unknown.map(function(x){return '{{'+x+'}}';}).join('、'));
  info.textContent=lines.join('\n');info.hidden=false;if(!d.used.length)say(t('模板里还没有占位符。请在要换的地方写上 {{列名}}（和名单第一行一样），例如：')+'{{'+d.columns[0]+'}}',true,true);else say('检查完成，可以生成了。');});}
function run(){return task(async function(){var tpl=templatePath(),list=listPath.value.trim();if(!list)throw Error('请先选择名单文件。');var mode=box.querySelector('input[name="mmMode"]:checked').value;say('正在生成…');
  var d=await api({template:tpl,list:list,mode:mode,nameColumn:mode==='separate'?nameSel.value:'',open:true});
  say(d.path?t('已生成 '+d.count+' 份，合成在一个文档里并已打开：')+d.path+(d.note?'\n'+t(d.note):''):t('已生成 '+d.count+' 个文件，放在：')+d.folder,false,true);});}
function build(){var files=document.querySelector('section[data-panel="files"]');if(!files||!T||T.host()!=='wps')return;
  box=el('section','tool-card mailmerge');box.id='mailMergeSection';box.dataset.panel='files';box.hidden=files.hidden;
  var head=el('div','card-heading'),icon=el('span','section-icon teal'),img=el('img');img.src='icons/duplicate.svg';img.alt='';icon.append(img);var h=el('div');h.append(el('h2','','按名单批量生成'),el('p','','邀请函、证书、通知书：一份模板 + 一张名单，每人一份'));head.append(icon,h);
  var steps=el('ol','mm-steps');['把当前文档当模板：要换的地方写 {{列名}}，例如 {{姓名}}（和名单第一行的列名一样），然后保存。','选择名单（Excel 表格，第一行是列名）。','点“检查”看对得上没有，再点“生成”。'].forEach(function(s){steps.append(el('li','',s));});
  var row=el('div','mm-row');listPath=el('input');listPath.type='text';listPath.placeholder=t('名单文件位置，例如 C:\\名单.xlsx');listPath.setAttribute('aria-label',t('名单文件位置'));var pick=el('button','','选择名单…');pick.type='button';pick.onclick=function(){var p=pickList();if(p===undefined){say('这里打不开选择窗口，请把名单的完整位置粘贴到左边的框里。',true);return;}if(p){listPath.value=p;check();}};row.append(listPath,pick);
  var modes=el('fieldset','mm-modes');modes.append(el('legend','','生成方式'));[['combined','合成一个文档（每份分页，方便一次打印）'],['separate','每行一个文件，文件名用：']].forEach(function(m,i){var l=el('label'),r=el('input');r.type='radio';r.name='mmMode';r.value=m[0];r.checked=!i;l.append(r,el('span','',m[1]));modes.append(l);});
  nameSel=el('select');nameSel.setAttribute('aria-label',t('文件名用哪一列'));fillColumns([]);modes.append(nameSel);
  info=el('pre','mm-info');info.hidden=true;info.setAttribute('translate','no');status=el('p','mm-status');status.setAttribute('role','status');status.setAttribute('translate','no');
  var bar=el('div','button-pair'),chk=el('button','','检查'),go=el('button','primary','生成');chk.type=go.type='button';chk.onclick=check;go.onclick=run;bar.append(chk,go);
  var tip=el('details','help-details');tip.append(el('summary','','占位符怎么写？'),el('p','','{{姓名}}、{姓名}、《姓名》、«姓名» 都可以，括号里写名单的列名。只替换名单里有的列名，正文里其他的《书名》不会动。日期、金额按 Excel 里显示的样子填进去。生成的文件放在“生成文件”文件夹，原文档和名单都不改。'));
  box.append(head,steps,row,modes,bar,info,status,tip);var paste=document.getElementById('pasteBarSection');files.parentNode.insertBefore(box,files);if(paste&&paste.nextSibling!==box)paste.after(box);}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',build);else build();
root.MailMergePanel={check:check,run:run};
})(window);
