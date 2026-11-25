// 文件与图片：图片转文字（粘贴截图或选图片，也能识别当前文档里的全部图片）、PDF 转 Word、提取 / 压缩当前文档的图片、历史版本。
// 处理都在本机服务里完成（系统自带的文字识别和图片工具），不上传到网上。原文件不改，结果另存为新文件。
(function(root){
 'use strict';
 var T=root.OfficeTools,P=root.AiTextPanel;if(!T||!P)return;
 var section=document.createElement('section');section.className='tool-card file-tools-panel';section.dataset.panel='quick';section.id='fileToolsPanel';
 section.innerHTML='<details><summary><strong>文件与图片</strong><span>图片转文字 · PDF 转 Word · 提取 / 压缩图片 · 历史版本</span></summary>'
  +'<h4 id="ftOcrHead">图片转文字</h4>'
  +'<div id="ftPaste" class="ft-paste" tabindex="0" role="button" aria-describedby="ftPasteHint">在这里按 Ctrl+V 粘贴截图</div>'
  +'<p class="input-hint" id="ftPasteHint">截图：Windows 按 Win+Shift+S，Mac 按 Shift+Command+Control+4。也可以选择图片文件（最多 20 张）。</p>'
  +'<div class="button-pair"><button type="button" id="ftPick">选择图片…</button><button type="button" id="ftDocOcr">识别当前文档里的图片</button></div>'
  +'<input type="file" id="ftFiles" accept="image/png,image/jpeg,image/gif,image/bmp,image/tiff" multiple hidden>'
  +'<div id="ftThumbs" class="ft-thumbs"></div>'
  +'<label for="ftLang">文字语言</label><select id="ftLang"><option value="zh">中文</option><option value="ja">日文</option><option value="en">英文</option><option value="auto">按系统语言</option></select>'
  +'<button type="button" id="ftOcrGo" class="primary" disabled>识别文字</button>'
  +'<textarea id="ftText" rows="6" hidden translate="no" aria-label="识别结果"></textarea>'
  +'<div class="button-pair" id="ftTextActions" hidden><button type="button" id="ftCopy">复制</button><button type="button" id="ftInsert">插入文档</button></div>'
  +'<h4 id="ftPdfHead">PDF 转 Word</h4><p class="input-hint">文字、表格和图片转成可以编辑的 Word（画了框线的表格还原成表格）；扫描件会先做文字识别（最多 30 页）。新文件放在 PDF 旁边。</p>'
  +'<button type="button" id="ftPdf">选择 PDF 并转换…</button><input type="file" id="ftPdfFile" accept="application/pdf,.pdf" hidden>'
  +'<h4 id="ftImgHead">当前文档的图片</h4><p class="input-hint">读取的是最近一次保存的文件（docx、xlsx、pptx）。</p>'
  +'<div class="button-pair"><button type="button" id="ftExtract">提取全部图片</button><button type="button" id="ftCompress">压缩图片另存</button></div>'
  +'<label for="ftLevel">压缩程度</label><select id="ftLevel"><option value="75,1920">标准（清晰，适合打印）</option><option value="60,1280">更小（适合发邮件、微信）</option><option value="85,2560">轻度（几乎看不出变化）</option></select>'
  +'<h4 id="ftHistHead">历史版本</h4><p class="input-hint">“备份当前文件”和自动备份留下的副本。打开的是副本，原文件不会被替换。</p>'
  +'<button type="button" id="ftHistory">查看当前文件的历史版本</button><div id="ftHistList" class="ft-history"></div>'
  +'<div id="ftStatus" role="status" aria-live="polite"></div><div class="button-pair" id="ftResult" hidden><button type="button" id="ftOpen">打开</button><button type="button" id="ftReveal">打开所在文件夹</button></div>'
  +'</details>';
 var $=function(id){return section.querySelector('#ft'+id);};
 function place(){var anchor=document.getElementById('aiChatPanel')||document.getElementById('aiTextPanel');if(anchor&&anchor.parentNode){anchor.parentNode.insertBefore(section,anchor.nextSibling);return true;}return false;}
 if(!place())document.addEventListener('DOMContentLoaded',place);
 function lang(){return root.OfficeI18n&&root.OfficeI18n.lang&&root.OfficeI18n.lang()==='ja'?'ja':'zh';}
 $('Lang').value=lang();
 var images=[],busy=false,lastPath='';
 function status(s,err){$('Status').textContent=s||'';$('Status').className='fa-status'+(err?' error':'');}
 function result(p){lastPath=p||'';$('Result').hidden=!p;$('Open').hidden=!p||!canOpen(p);}
 var token='';
 // 服务重启后旧口令会失效（403）：重新取一次口令再试，用户不用重开 WPS。
 async function api(url,body,retry){if(!token){var s=await fetch('/api/session').then(function(r){return r.json();});token=s.token;}
  var r;try{r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify(body)});}catch(e){throw Error('连不上本机服务：请双击桌面“Office Toolbox”后再试。');}if(r.status===403&&!retry){token='';return api(url,body,true);}var d=await r.json().catch(function(){return {};});if(r.status===404&&!d.error)throw Error('本机服务还是旧版本：请双击桌面“Office Toolbox”重启后再试。');if(!r.ok)throw Error(d.error||'没有完成处理，请重试。');return d;}
 async function task(fn){if(busy)return;busy=true;section.querySelectorAll('button').forEach(function(b){b.dataset.was=b.disabled?'1':'';b.disabled=true;});try{await fn();}catch(e){status(e.message,true);}finally{busy=false;section.querySelectorAll('button').forEach(function(b){b.disabled=b.dataset.was==='1';});$('OcrGo').disabled=!images.length;}}
 function mb(n){return n>=1048576?(n/1048576).toFixed(1)+' MB':Math.max(1,Math.round(n/1024))+' KB';}
 function readFile(f){return new Promise(function(resolve,reject){var fr=new FileReader();fr.onload=function(){resolve(String(fr.result));};fr.onerror=function(){reject(Error('文件读取失败。'));};fr.readAsDataURL(f);});}
 // ---- 当前文档（WPS：直接读路径并能打开；Microsoft Office：只能读路径，打开要手动） ----
 // Microsoft Office 侧栏里 OfficeTools.app() 是模拟对象（office-shim 会设置 AiTextHost），不能当 WPS 用。
 function wps(){if(root.AiTextHost)return null;try{return T&&T.app&&T.app();}catch(e){return null;}}
 async function current(){var a=wps();if(a){var h=T.host(),d=h==='wps'?a.ActiveDocument:h==='et'?a.ActiveWorkbook:a.ActivePresentation;if(!d)throw Error('请先打开一份文档。');var p=String(d.FullName||'');if(!/[\\\/]/.test(p))throw Error('这份文档还没有保存到电脑上，请先按 Ctrl+S 保存。');return {path:p,saved:d.Saved!==false&&d.Saved!==0};}
  var helper=root.OfficeAddinTools,u=helper&&helper.documentUrl?await helper.documentUrl({Office:root.Office}):root.Office&&Office.context&&Office.context.document&&Office.context.document.url;
  var p=helper&&helper.localPath?helper.localPath(u):u&&!/^https?:/i.test(u)?u:'';if(p)return {path:p,saved:true};if(u)throw Error('这份文档存在网上（OneDrive / SharePoint），请先另存到电脑上。');throw Error('这份文档还没有保存到电脑上，请先保存。');}
 async function needSaved(){var c=await current();if(!c.saved&&!(await T.lib.ask('文档有还没保存的修改，工具读取的是上次保存的内容。是否继续？（想包含最新修改，请先按 Ctrl+S）')))return null;return c;}
 function canOpen(p){var a=wps();if(!a)return false;var h=T.host(),ext=(p.match(/\.[^.\\\/]+$/)||[''])[0].toLowerCase();return (h==='wps'&&/^\.(docx?|wps|rtf|txt)$/.test(ext))||(h==='et'&&/^\.(xlsx?|et|csv)$/.test(ext))||(h==='wpp'&&/^\.(pptx?|dps)$/.test(ext));}
 function openFile(p){var a=wps(),h=T.host();if(h==='wps')a.Documents.Open(p);else if(h==='et')a.Workbooks.Open(p);else a.Presentations.Open(p);}
 // WPS 里用自己的“打开文件”对话框选 PDF（结果放在 PDF 旁边）；不行就用网页的文件选择（结果放到备份文件夹）。
 function pickPath(title,desc,pattern){var a=wps();if(!a||typeof a.FileDialog!=='function')return undefined;try{var dlg=a.FileDialog(3),tr=root.OfficeI18n&&root.OfficeI18n.t;if(tr){title=tr(title);desc=tr(desc);}dlg.Title=title;dlg.AllowMultiSelect=false;try{dlg.Filters.Clear();dlg.Filters.Add(desc,pattern);}catch(ignore){}if(dlg.Show()!==-1)return null;return String(dlg.SelectedItems.Item(1));}catch(e){return undefined;}}
 // ---- 图片转文字 ----
 function addImage(dataUrl,name){if(images.length>=20){status('一次最多 20 张图片。',true);return;}images.push({data:dataUrl,name:name||(lang()==='ja'?'スクリーンショット':'截图')});drawThumbs();}
 function drawThumbs(){var box=$('Thumbs');box.textContent='';images.forEach(function(im,i){var fig=document.createElement('figure'),img=document.createElement('img'),x=document.createElement('button');img.src=im.data;img.alt=im.name;x.type='button';x.textContent='×';x.title='去掉这张';x.onclick=function(){images.splice(i,1);drawThumbs();};fig.append(img,x);box.append(fig);});$('OcrGo').disabled=!images.length;if(images.length)status('已添加 '+images.length+' 张图片，点“识别文字”。');}
 function onPaste(e){var items=(e.clipboardData&&e.clipboardData.items)||[],got=false;for(var i=0;i<items.length;i++){var it=items[i];if(it.kind==='file'&&/^image\//.test(it.type)){var f=it.getAsFile();if(f){got=true;readFile(f).then(function(u){addImage(u,'截图');});}}}if(got)e.preventDefault();else if(e.target===$('Paste'))status('剪贴板里没有图片。请先截图，再在这里按 Ctrl+V。',true);}
 $('Paste').addEventListener('paste',onPaste);
 section.addEventListener('paste',function(e){if(e.target!==$('Paste')&&e.target!==$('Text'))onPaste(e);});
 $('Pick').onclick=function(){$('Files').click();};
 $('Files').onchange=function(){var fs=[].slice.call($('Files').files||[]);$('Files').value='';fs.forEach(function(f){if(f.size>20*1024*1024){status(f.name+' 超过 20 MB，请缩小后再试。',true);return;}readFile(f).then(function(u){addImage(u,f.name);});});};
 function showText(d){$('Text').hidden=false;$('TextActions').hidden=false;$('Text').value=d.text;status('识别完成（'+d.engine+' 文字识别）。请核对后再使用，识别可能有错字。'+(d.note?'\n'+d.note:''));}
 $('OcrGo').onclick=function(){task(async function(){result('');status('正在识别，通常需要几秒…');var d=await api('/api/tools/ocr',{images:images.map(function(x){return x.data;}),lang:$('Lang').value});showText(d);});};
 $('DocOcr').onclick=function(){task(async function(){result('');var c=await needSaved();if(!c)return status('已取消。');status('正在识别文档里的图片…');var d=await api('/api/tools/ocr',{path:c.path,lang:$('Lang').value});showText(d);status($('Status').textContent+'\n共 '+d.count+' 张图片。');});};
 $('Copy').onclick=function(){var t=$('Text');t.select();var ok=false;try{ok=document.execCommand('copy');}catch(e){}if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(t.value).then(function(){status('已复制。');},function(){status(ok?'已复制。':'复制失败，请选中文字后按 Ctrl+C。',!ok);});else status(ok?'已复制。':'复制失败，请选中文字后按 Ctrl+C。',!ok);};
 $('Insert').onclick=function(){task(async function(){var host=P.host();if(!host.insert)throw Error('这个程序暂时不支持插入，请复制后粘贴。');var text=$('Text').value,r=await host.insert(text,false);if(r&&r.confirm){if(!(await T.lib.ask(r.confirm)))return status('已取消。');r=await host.insert(text,true);}status(r);});};
 // ---- PDF 转 Word ----
 async function convert(body){status('正在转换，页数多或扫描件需要一两分钟…');var d=await api('/api/tools/pdf-to-word',Object.assign({lang:$('Lang').value},body));
  status(['已转换为 Word：'+d.path,'共 '+d.pages+' 页、'+d.paragraphs+' 段。'].concat(d.tables?['还原了 '+d.tables+' 个表格。']:[],d.images?['带过去 '+d.images+' 张图片。']:[],d.ocr?['扫描件，识别了 '+d.ocr+' 页，请核对错字和表格。']:[],[d.ocr?'扫描件的表格按格子位置还原，合并单元格和很淡的框线可能认不准，请对照原件检查。':'没有框线的表格会变成一行一行的文字，版式需要再整理。'],d.notes||[]).join('\n'));result(d.path);}
 $('Pdf').onclick=function(){var p=pickPath('选择要转换的 PDF','PDF 文件','*.pdf');if(p===null)return;if(p)return task(function(){return convert({path:p});});$('PdfFile').click();};
 $('PdfFile').onchange=function(){var f=$('PdfFile').files&&$('PdfFile').files[0];$('PdfFile').value='';if(!f)return;if(f.size>80*1024*1024)return status('PDF 超过 80 MB，请先拆分。',true);task(async function(){var u=await readFile(f);await convert({data:u,name:f.name});});};
 // ---- 提取 / 压缩图片 ----
 $('Extract').onclick=function(){task(async function(){result('');var c=await needSaved();if(!c)return status('已取消。');var d=await api('/api/tools/extract-images',{path:c.path});status('已把 '+d.count+' 张图片存到：'+d.dir+'\n文档本身没有改动。');result(d.dir);});};
 $('Compress').onclick=function(){task(async function(){result('');var c=await needSaved();if(!c)return status('已取消。');var lv=$('Level').value.split(',');status('正在压缩图片…');var d=await api('/api/tools/compress-images',{path:c.path,quality:Number(lv[0]),maxSide:Number(lv[1])});
  if(!d.changed)return status('文档里的 '+d.count+' 张图片已经足够小，没有需要压缩的。');status('已压缩 '+d.changed+' 张图片（共 '+d.count+' 张）：图片 '+mb(d.before)+' → '+mb(d.after)+'，文件 '+mb(d.fileBefore)+' → '+mb(d.fileAfter)+'。\n另存为：'+d.path+'\n原文件没有改动。');result(d.path);});};
 // ---- 历史版本 ----
 $('History').onclick=function(){task(async function(){result('');var c=await current(),d=await api('/api/tools/history',{path:c.path}),box=$('HistList');box.textContent='';
  if(!d.items.length){status('还没有这个文件的备份。可以用“备份当前文件”，或在设置里打开自动备份。');return;}
  d.items.forEach(function(it){var row=document.createElement('div'),info=document.createElement('span'),open=document.createElement('button');row.className='ft-hist-row';info.textContent=it.time+(it.auto?'（自动）':'')+' · '+mb(it.size);open.type='button';open.textContent=canOpen(it.path)?'打开这个版本':'打开所在文件夹';
   open.onclick=function(){try{if(canOpen(it.path)){openFile(it.path);status('已打开 '+it.time+' 的版本（副本）。需要恢复时，请把这个副本“另存为”原来的文件名。');}else api('/api/tools/reveal',{path:it.path}).catch(function(e){status(e.message,true);});}catch(e){status(e.message,true);}};row.append(info,open);box.append(row);});
  status('找到 '+d.items.length+' 个版本（新的在前），存放在：'+d.dir);});};
 $('Open').onclick=function(){if(!lastPath)return;try{openFile(lastPath);status('已打开：'+lastPath);}catch(e){status('没能打开：'+e.message,true);}};
 $('Reveal').onclick=function(){if(lastPath)api('/api/tools/reveal',{path:lastPath}).catch(function(e){status(e.message,true);});};
 function select(action){var head={ocr:'OcrHead',pdf:'PdfHead',images:'ImgHead',history:'HistHead'}[action];var dt=section.querySelector('details');if(dt)dt.open=true;if(head){var h=$(head);if(h&&h.scrollIntoView)h.scrollIntoView({block:'start'});}if(action==='ocr')$('Paste').focus();return !!head;}
 root.FileToolsPanel={select:select,images:function(){return images.slice();},api:api};
})(window);
