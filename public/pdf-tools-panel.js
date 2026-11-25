(function(root){'use strict';
// PDF 工具（侧栏“文件与内容”页，WPS 文字 / 表格 / 演示里都显示）：合并、提取 / 删除页、拆分、PDF 转图片、图片转 PDF。
// 选好文件、点“开始”，结果放进“生成文件”文件夹并在文件夹里选中；原文件不改。
var T=root.OfficeTools,box=null,files=null,status=null,info=null,act=null,busy=false,opts={};
function t(s){return root.OfficeI18n?root.OfficeI18n.t(s):s;}
function el(tag,cls,text){var e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;}
function say(m,bad,raw){status.textContent=raw?m:t(m);status.className='mm-status'+(bad?' bad':'');}
async function api(body){var L=T.lib,s=await L.localFetch('/api/session',{}).then(function(r){if(!r.ok)throw Error('本地服务未连接。');return r.json();});var r=await L.localFetch('/api/tools/pdf',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(body)});var d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(d.error||'没有完成，请重试。');return d;}
var ACTIONS=[
  {id:'merge',label:'合并几个 PDF',hint:'按列表里的顺序接成一个 PDF。每行一个文件，可以上下调整行的顺序。',pdf:true,multi:true},
  {id:'pick',label:'提取部分页面',hint:'把选中的页另存成新的 PDF。',pdf:true,pages:'要提取的页，例如：1-3,5,8-'},
  {id:'remove',label:'删除部分页面',hint:'删掉这些页，其余的另存成新的 PDF。',pdf:true,pages:'要删除的页，例如：2,4-5'},
  {id:'split',label:'拆分成几个 PDF',hint:'每几页存一个文件，或者按段拆（例如 1-3,4-10,11-）。',pdf:true,split:true},
  {id:'to-images',label:'PDF 转图片',hint:'每页存成一张 PNG 图片（不填页码 = 全部页，最多 100 页）。',pdf:true,pages:'只转这些页（可不填），例如：1-3',dpi:true},
  {id:'images-to-pdf',label:'图片转 PDF',hint:'每张图片一页（JPG、PNG），按列表里的顺序。',multi:true,size:true}];
function current(){var id=act.value;return ACTIONS.find(function(a){return a.id===id;})||ACTIONS[0];}
function list(){return files.value.split(/\r?\n/).map(function(s){return s.trim().replace(/^"(.*)"$/,'$1');}).filter(Boolean);}
function pick(){var a=current();try{var app=T.app(),dlg=app.FileDialog(3);dlg.Title=t(a.pdf?'选择 PDF 文件':'选择图片');dlg.AllowMultiSelect=!!a.multi;try{dlg.Filters.Clear();if(a.pdf)dlg.Filters.Add('PDF','*.pdf');else dlg.Filters.Add(t('图片'),'*.jpg;*.jpeg;*.png');}catch(ignore){}
  if(dlg.Show()!==-1)return null;var out=[];for(var i=1;i<=Number(dlg.SelectedItems.Count||0);i++)out.push(String(dlg.SelectedItems.Item(i)));return out;}catch(e){return undefined;}}
function refresh(){var a=current();box.querySelector('.pdf-hint').textContent=t(a.hint);opts.pages.hidden=!a.pages;if(a.pages)opts.pagesInput.placeholder=t(a.pages);opts.split.hidden=!a.split;opts.dpi.hidden=!a.dpi;opts.size.hidden=!a.size;
  files.placeholder=t(a.pdf?(a.multi?'PDF 文件位置，每行一个，例如 C:\\资料\\合同.pdf':'PDF 文件位置，例如 C:\\资料\\合同.pdf'):'图片位置，每行一个，例如 C:\\资料\\扫描1.jpg');files.rows=a.multi?4:2;info.textContent='';}
function showInfo(){var a=current(),l=list();if(!a.pdf||a.multi||l.length!==1){info.textContent='';return;}api({action:'info',path:l[0]}).then(function(d){info.textContent=t('共 '+d.pages+' 页')+'：'+d.name;}).catch(function(){info.textContent='';});}
function run(){if(busy)return;var a=current(),l=list(),body={action:a.id==='remove'?'pick':a.id,open:true};
  if(!l.length){say(a.pdf?'请先选择 PDF 文件。':'请先选择图片。',true);return;}if(a.multi)body.paths=l;else{if(l.length>1){say('这一项一次处理一个 PDF，请只留一行。',true);return;}body.path=l[0];}
  if(a.id==='remove')body.mode='remove';if(a.pages)body.pages=opts.pagesInput.value;if(a.split){body.every=opts.everyInput.value;body.ranges=opts.rangesInput.value;}if(a.dpi)body.dpi=Number(opts.dpiSelect.value);if(a.size)body.size=opts.sizeSelect.value;
  busy=true;box.querySelectorAll('button').forEach(function(b){b.disabled=true;});say('正在处理…');
  api(body).then(function(d){if(d.path)say(t(d.files?'已合并 '+d.files+' 个文件，共 '+d.pages+' 页，已在文件夹里选中：':'已生成新的 PDF（'+d.pages+' 页），已在文件夹里选中：')+d.path,false,true);
    else say(t(d.dpi?'已转成 '+d.count+' 张图片，放在：':'已拆成 '+d.count+' 个文件，放在：')+d.folder,false,true);})
  .catch(function(e){say(t('未完成：')+t(e.message),true,true);}).then(function(){busy=false;box.querySelectorAll('button').forEach(function(b){b.disabled=false;});});}
function field(labelText,input){var l=el('label','pdf-field');l.append(el('span','',labelText),input);return l;}
function select(pairs,value){var s=el('select');pairs.forEach(function(p){var o=el('option','',p[1]);o.value=p[0];s.append(o);});s.value=value;return s;}
function build(){var panel=document.querySelector('section[data-panel="files"]');if(!panel||!T)return;
  box=el('section','tool-card mailmerge pdftools');box.id='pdfToolsSection';box.dataset.panel='files';box.hidden=panel.hidden;
  var head=el('div','card-heading'),icon=el('span','section-icon teal'),img=el('img');img.src='icons/pdf.svg';img.alt='';icon.append(img);var h=el('div');h.append(el('h2','','PDF 工具'),el('p','','合并、拆分、提取页面，PDF 和图片互转（原文件不改）'));head.append(icon,h);
  act=select(ACTIONS.map(function(a){return [a.id,a.label];}),'merge');act.id='pdfAction';act.setAttribute('aria-label',t('要做什么'));act.onchange=function(){refresh();showInfo();};
  var hint=el('p','pdf-hint');files=el('textarea','pdf-files');files.setAttribute('aria-label',t('文件位置'));files.spellcheck=false;files.onchange=showInfo;
  var row=el('div','mm-row'),choose=el('button','','选择文件…');choose.type='button';choose.onclick=function(){var p=pick();if(p===undefined){say('这里打不开选择窗口，请把文件的完整位置粘贴到上面的框里（每行一个）。',true);return;}if(!p||!p.length)return;var a=current();files.value=(a.multi&&files.value.trim()?files.value.trim()+'\n':'')+p.join('\n');showInfo();};
  var clear=el('button','','清空');clear.type='button';clear.onclick=function(){files.value='';info.textContent='';};row.append(choose,clear);info=el('p','mm-info pdf-info');info.setAttribute('translate','no');
  opts.pagesInput=el('input');opts.pagesInput.type='text';opts.pages=field('页码',opts.pagesInput);
  opts.everyInput=el('input');opts.everyInput.type='number';opts.everyInput.min='1';opts.everyInput.value='1';opts.rangesInput=el('input');opts.rangesInput.type='text';opts.rangesInput.placeholder=t('或者按段拆，例如：1-3,4-10,11-');
  opts.split=el('div','pdf-split');opts.split.append(field('每几页一个文件',opts.everyInput),field('按段拆（填了就不按页数拆）',opts.rangesInput));
  opts.dpiSelect=select([['150','一般（150 dpi）'],['200','清晰（200 dpi，文件更大）']],'150');opts.dpi=field('清晰度',opts.dpiSelect);
  opts.sizeSelect=select([['A4','放进 A4 纸（横图自动横放）'],['fit','页面和图片一样大']],'A4');opts.size=field('页面大小',opts.sizeSelect);
  status=el('p','mm-status');status.setAttribute('role','status');status.setAttribute('translate','no');
  var bar=el('button','primary','开始');bar.type='button';bar.onclick=run;
  box.append(head,field('要做什么',act),hint,files,row,info,opts.pages,opts.split,opts.dpi,opts.size,bar,status);
  var after=document.getElementById('mailMergeSection')||document.getElementById('pasteBarSection');if(after)after.after(box);else panel.parentNode.insertBefore(box,panel);refresh();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',build);else build();
root.PdfToolsPanel={run:run};
})(window);
