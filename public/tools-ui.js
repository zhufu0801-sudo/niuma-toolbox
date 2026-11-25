(function(){
  'use strict';
  var h=OfficeTools.host(),busy=false,$=function(id){return document.getElementById(id);},I18n=window.OfficeI18n||null;
  function t(text){return I18n?I18n.t(text):String(text);}
  function mark(text){return I18n?I18n.mark(text):String(text);}
  var hostClass={et:'teal',wpp:'purple'};
  function status(text,error){$('status').textContent=text;$('statusBox').className='status-box '+(error?'error':busy?'':'success');$('statusIcon').src=ToolIcons.path(error?'clean':'check');}
  // 换了文档（WPS 新窗口会沿用同一个侧栏）时，清掉上一个文档的操作结果，免得点“打开”打开别的文件。
  var lastDocKey='';
  function clearOldResults(){try{document.querySelectorAll('#libraryMessage,#ftStatus,#status,#statusBox,.fa-status,.ac-suggest,[data-result]').forEach(function(el){if(el.classList&&el.classList.contains('ac-suggest')){el.remove();return;}if(el.id==='status'||el.id==='statusBox')return;el.textContent='';});var r=document.getElementById('ftResult');if(r)r.hidden=true;}catch(ignore){}}
  function refresh(){try{var a=OfficeTools.app(),d=h==='wps'?a.ActiveDocument:h==='et'?a.ActiveWorkbook:a.ActivePresentation;var name=d.Name,full='';try{full=String(d.FullName||name);}catch(ignore){full=name;}if(lastDocKey&&lastDocKey!==full)clearOldResults();lastDocKey=full;$('document').textContent=name;$('document').title=name;$('hostLabel').textContent=({wps:'WPS 文字',et:'WPS 表格',wpp:'WPS 演示'}[h])+(d.ReadOnly?' · 只读':' · 当前文档');$('docDot').className='connection-dot connected';}catch(e){$('hostLabel').textContent='界面预览';$('document').textContent=I18n?I18n.t('在 WPS 中打开文档即可使用'):'在 WPS 中打开文档即可使用';$('document').title='请从 WPS 顶部“办公工具箱”打开';$('docDot').className='connection-dot';}}
  async function run(id,extra){if(busy)return;busy=true;document.querySelectorAll('button,textarea,input').forEach(function(b){b.disabled=true;});status('正在执行，请稍候…');try{var options={host:h,text:$('content').value,findText:$('findText').value,replaceText:$('replaceText').value};if(extra)Object.keys(extra).forEach(function(k){options[k]=extra[k];});var result=await OfficeTools.runFromUi(id,options);busy=false;status(result);}catch(e){busy=false;status(t('未完成：')+t(e.message)+t(h==='et'?'\n如已出现部分更改，请检查表格内容；工具写入的内容可能无法用撤销恢复。':'\n如已出现部分更改，可使用 WPS 自带撤销。'),true);}finally{busy=false;document.querySelectorAll('button,textarea,input').forEach(function(b){b.disabled=false;});refresh();if(extra&&extra.after)extra.after();}}
  function setView(view){if(['quick','create','cards','files'].indexOf(view)<0)view='quick';if(view==='cards')renderCards();document.querySelectorAll('[data-view]').forEach(function(b){b.classList.toggle('active',b.dataset.view===view);b.setAttribute('aria-pressed',String(b.dataset.view===view));});document.querySelectorAll('[data-panel]').forEach(function(panel){panel.hidden=panel.dataset.panel!==view||(panel.id==='creation'&&h==='et')||(panel.id==='sheetCreationNote'&&h!=='et')||(panel.id==='replaceSection'&&h==='wpp');});$('pageTitle').textContent={quick:'常用工具',create:h==='wpp'?'大纲制作':'内容制作',cards:'内容卡片',files:'文件与内容'}[view];$('pageSubtitle').textContent={quick:'少找一次菜单，多一点从容。',create:'从一个想法，到一份清晰的内容。',cards:'选好内容设为模块，点插入放进其他文档。',files:'让整理好的内容，自由流转。'}[view];}
  // 常用列表来自功能登记表；用户在“选择要显示的功能”里关掉的不显示。
  function makeTool(item){var button=document.createElement('button');button.type='button';button.className='quick-tool';button.title=item.desc||'';var icon=document.createElement('span'),img=document.createElement('img'),title=document.createElement('span'),desc=document.createElement('span');icon.className='tool-icon';if(hostClass[h])icon.classList.add(hostClass[h]);img.src=ToolIcons.path(item.icon);img.alt='';icon.append(img);title.className='tool-title';title.textContent=item.label;desc.className='tool-desc';desc.textContent=item.desc||'';button.append(icon,title,desc);button.onclick=function(){run(item.id);};return button;}
  function renderQuick(){var box=$('quick'),items=OfficeTools.list(h).filter(function(x){return x.panel&&!OfficeTools.isHidden(x.id);});box.textContent='';items.forEach(function(item){box.append(makeTool(item));});$('quickEmpty').hidden=items.length>0;}
  function renderManage(){var box=$('manageList'),hide=OfficeTools.hidden();box.textContent='';OfficeTools.list(h).filter(function(x){return x.label&&x.id!=='insertCard'&&x.id!=='insertLinkedModule';}).forEach(function(item){var label=document.createElement('label'),box2=document.createElement('input'),text=document.createElement('span');label.className='manage-item';box2.type='checkbox';box2.checked=hide.indexOf(item.id)<0;box2.onchange=async function(){var ids=OfficeTools.hidden().filter(function(id){return id!==item.id;});if(!box2.checked)ids.push(item.id);box2.disabled=true;try{await OfficeTools.setHidden(ids);renderQuick();status(box2.checked?'已显示“'+item.label+'”。':'已隐藏“'+item.label+'”，顶部功能区稍后同步。');}catch(e){status('设置未保存：'+e.message,true);}finally{renderManage();}};text.textContent=item.label;label.append(box2,text);box.append(label);});}
  renderQuick();renderManage();if(window.WorkbenchSettings){WorkbenchSettings.subscribe(function(){renderQuick();renderManage();});OfficeTools.ready().catch(function(e){status('设置暂时无法读取：'+e.message,true);});}var versionBadge=$('versionBadge');if(versionBadge)versionBadge.textContent='v'+OfficeTools.version;
  document.querySelectorAll('[data-action]').forEach(function(b){b.onclick=function(){run(b.dataset.action);};});
  document.querySelectorAll('[data-view]').forEach(function(b){b.onclick=function(){setView(b.dataset.view);};});document.querySelectorAll('[data-jump]').forEach(function(b){b.onclick=function(){setView(b.dataset.jump);};});
  function studioNotice(result){var note=$('studioNotice'),link=$('studioFallback');note.hidden=false;note.textContent=result.message;note.className='inline-notice'+(result.ok?'':' error');link.hidden=false;link.href=result.url;}
  // 侧栏 WebView 不一定能打开新窗口，优先用 WPS 对话框打开设置页，失败时再走普通链接。
  $('settingsEntry').onclick=function(e){var a=OfficeTools.app();try{if(a&&typeof a.ShowDialog==='function'){a.ShowDialog(location.origin+'/settings.html',I18n?I18n.t('办公工具箱设置'):'办公工具箱设置',760,760,false);if(e&&e.preventDefault)e.preventDefault();return false;}}catch(ignore){}return true;};
  $('studioLaunch').onclick=function(){studioNotice(StudioLaunch.open({app:OfficeTools.app(),origin:location.origin,open:window.open.bind(window)}));};
  $('documentIcon').src=ToolIcons.path(h==='et'?'table':h==='wpp'?'presentation':'document');
  if(h==='wpp'){$('creationTitle').textContent='按大纲制作 PPT';$('creationIcon').src=ToolIcons.path('presentation');$('inputHint').textContent='每页第一行作标题，其余行作正文；用空行分隔不同页。一次最多 30 页。';$('content').placeholder='项目概览\n项目背景\n本期目标\n\n工作计划\n第一阶段任务\n第二阶段任务';$('createLabel').textContent='在当前页之后生成页面';$('creationNote').textContent='插在当前页之后，不改动现有页面';}
  $('create').onclick=function(){run(h==='wpp'?'makeSlides':'insertTable');};$('replaceButton').onclick=function(){run('replaceText');};
  $('content').oninput=function(){var text=$('content').value.trim();$('preview').textContent=!text?'':h==='wpp'?'将生成 '+text.split(/\n\s*\n/).length+' 页；第一行标题，其余行正文。':'将插入 '+text.split('\n').length+' 行表格，保留原文。';};
  // ---- 内容卡片 ----
  var SOURCE={wps:'文字',et:'表格',wpp:'演示',drop:'拖入'},SOURCE_JA={wps:'Writer',et:'Spreadsheets',wpp:'Presentation',drop:'ドラッグ'};
  function sourceName(x){return I18n&&I18n.lang()==='ja'?(SOURCE_JA[x]||''):(SOURCE[x]||'');}
  var cardsView=false,cardsSignature='',cardsRequest=0,cardDragging=false;
  function plainOf(b){return b.type==='table'?b.rows.map(function(r){return r.join('\t');}).join('\n'):b.text;}
  function esc(t){return String(t).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  function htmlOf(b){return b.type==='table'?'<table border="1">'+b.rows.map(function(r){return '<tr>'+r.map(function(c){return '<td>'+esc(c)+'</td>';}).join('')+'</tr>';}).join('')+'</table>':'<p>'+esc(b.text).replace(/\n/g,'<br>')+'</p>';}
  function readDrop(dt){
    var html='',text='';try{html=dt.getData('text/html')||'';}catch(ignore){}try{text=dt.getData('text/plain')||dt.getData('Text')||'';}catch(ignore){}
    var cut=false;
    if(html&&typeof DOMParser==='function'){var table=new DOMParser().parseFromString(html,'text/html').querySelector('table');if(table){var rows=[].slice.call(table.rows).map(function(tr){return [].slice.call(tr.cells).map(function(td){return String(td.innerText||td.textContent||'').replace(/\s+$/,'').replace(/^\s+/,'');});});if(rows.length>20||rows.some(function(r){return r.length>10;}))cut=true;rows=rows.slice(0,20).map(function(r){return r.slice(0,10);});if(rows.length&&rows[0].length)return {block:{type:'table',rows:rows},cut:cut};}}
    text=text.replace(/\r\n?/g,'\n').replace(/\n+$/,'');if(!text.trim())return null;
    var lines=text.split('\n');if(lines.length>1&&lines.every(function(l){return l.indexOf('\t')>=0;})){var t=lines.map(function(l){return l.split('\t');});if(t.length>20||t.some(function(r){return r.length>10;}))cut=true;return {block:{type:'table',rows:t.slice(0,20).map(function(r){return r.slice(0,10);})},cut:cut};}
    if(text.length>10000){cut=true;text=text.slice(0,10000);}
    return {block:{type:'text',text:text},cut:cut};
  }
  function cardItem(card){
    var el=document.createElement('div'),title=document.createElement('div'),meta=document.createElement('div'),actions=document.createElement('div'),ins=document.createElement('button'),del=document.createElement('button');
    el.className='card-item';if(card.sourceHost)el.dataset.source=card.sourceHost;el.draggable=true;el.title='按住可以直接拖进文档';title.className='card-title';title.textContent=I18n?I18n.t(String(card.title).replace(/^(\d+ 行 × \d+ 列表格：)?([\s\S]*)$/,function(all,head,rest){return (head||'')+I18n.mark(rest);})):card.title;
    title.setAttribute('translate','no');
    meta.className='card-meta';meta.textContent=(card.block.type==='table'?'表格':'文字')+' · 来自'+mark(sourceName(card.sourceHost))+' · '+new Date(card.created).toLocaleString(I18n&&I18n.lang()==='ja'?'ja-JP':'zh-CN');
    actions.className='card-actions';ins.type='button';ins.className='insert';ins.textContent=card.source?'普通插入':'插入';del.type='button';del.textContent='删除';
    ins.onclick=function(){run('insertCard',{block:card.block,format:card.format});};
    del.onclick=function(){if(card.source&&!window.confirm('删除这个联动模块后，其目标内容仍保留，但无法继续更新。确定删除吗？'))return;OfficeTools.cards.remove(card.id).then(function(){status('已删除卡片。');renderCards();},function(e){status('未完成：'+e.message,true);});};
    el.ondragstart=function(e){cardDragging=true;try{e.dataTransfer.effectAllowed='copy';e.dataTransfer.setData('text/plain',plainOf(card.block));e.dataTransfer.setData('text/html',card.html||htmlOf(card.block));}catch(ignore){}};
    el.ondragend=function(){cardDragging=false;};
    var desktop=document.createElement('button');desktop.type='button';desktop.textContent='保存到桌面';desktop.title='按来源默认格式保存为新文件；格式可在设置里调整';desktop.onclick=async function(){if(busy)return;desktop.disabled=true;status('正在生成文件…');try{var result=await OfficeTools.cards.saveDesktop(card.id);status('已保存：'+result.filePath+(result.warnings&&result.warnings.length?'\n'+result.warnings.join(' '):''));}catch(e){status('保存未完成：'+e.message,true);}finally{desktop.disabled=false;}};
    actions.append(ins);if(card.source&&OfficeTools.links){var linked=document.createElement('button'),source=document.createElement('button');linked.type='button';linked.textContent='联动插入';linked.onclick=function(){run('insertLinkedModule',{cardId:card.id,after:renderCards});};source.type='button';source.textContent='刷新来源';source.onclick=async function(){if(busy)return;source.disabled=true;try{var updated=await OfficeTools.links.refreshSource(card);status('来源已刷新，第 '+(updated.revision||1)+' 版。请在目标点击“更新本文件联动”。');renderCards();}catch(e){status('来源未刷新：'+e.message,true);}finally{source.disabled=false;}};actions.append(linked);if(card.source.host===h)actions.append(source);meta.title=card.source.path;meta.textContent+=' · 联动来源 · 第 '+(card.revision||1)+' 版';var path='';try{var a=OfficeTools.app();path=String((h==='wps'?a.ActiveDocument:h==='et'?a.ActiveWorkbook:a.ActivePresentation).FullName).split(String.fromCharCode(92)).join('/').toLowerCase();}catch(ignore){}var local=(card.targets||[]).filter(function(link){return link.anchor.host===h&&String(link.anchor.path).split(String.fromCharCode(92)).join('/').toLowerCase()===path;});if(local.length){var detach=document.createElement('button');detach.type='button';detach.textContent='解除本文件联动';detach.onclick=async function(){if(!window.confirm('解除本文件中这个模块的联动？已插入内容保留，之后不再跟随来源更新。'))return;try{for(var i=0;i<local.length;i++)await OfficeTools.links.detach(card.id,local[i].id);status('已解除联动，文档内容保留。');renderCards();}catch(e){status('解除未完成：'+e.message,true);}};actions.append(detach);}}actions.append(desktop,del);el.append(title,meta,actions);return el;
  }
  function renderCards(){if(!OfficeTools.cards||busy||cardDragging)return;var request=++cardsRequest;OfficeTools.cards.list().then(function(list){if(request!==cardsRequest||busy||cardDragging)return;var signature=JSON.stringify(list);if(signature===cardsSignature)return;cardsSignature=signature;var box=$('cardList');box.textContent='';list.forEach(function(c){box.append(cardItem(c));});$('cardsEmpty').hidden=list.length>0;},function(e){if(request!==cardsRequest)return;$('cardsEmpty').hidden=false;$('cardsEmpty').textContent='卡片暂时读不出来：'+e.message;});}
  (function(){var zone=$('dropZone');if(!zone)return;
    zone.ondragover=function(e){e.preventDefault();try{e.dataTransfer.dropEffect='copy';}catch(ignore){}zone.classList.add('over');};
    zone.ondragleave=function(){zone.classList.remove('over');};
    zone.ondrop=function(e){e.preventDefault();zone.classList.remove('over');var got=readDrop(e.dataTransfer);if(!got){status('没有读到可以存的文字或表格，请先选中内容再拖进来。',true);return;}
      OfficeTools.cards.add(got.block,'drop').then(function(card){status('已存成卡片：'+mark(card.title)+(got.cut?'（内容较多，只保存了前面一部分）':''));renderCards();},function(err){status('未完成：'+err.message,true);});};
    $('saveCardButton').onclick=function(){run('saveCard',{after:renderCards});};
    // 表格里按住选区拖动，WPS 只会在表格内部移动单元格，内容不会进到侧栏；所以表格窗口直接提示用按钮。
    if(h==='wpp'&&$('undoNote'))$('undoNote').textContent='部分演示操作无法撤销，建议先保存副本。';if(h==='et'){if($('undoNote'))$('undoNote').textContent='表格里的修改不能用撤销还原，重要文件请先点“备份当前文件”';zone.classList.add('sheet-hint');zone.querySelector('strong').textContent='表格里拖动选区只会移动单元格';zone.querySelector('small').textContent='请先选中区域，再点下面的“把当前选区设为模块”';zone.setAttribute('aria-label','表格里请用下面的按钮把选区存成模块');$('saveCardButton').classList.add('emphasis');}
  })();
  setInterval(function(){if(!$('cardsSection').hidden)renderCards();},4000);
  OfficeTools.cardsUi={readDrop:readDrop,plainOf:plainOf,htmlOf:htmlOf};
  function showView(view){setView(view);try{history.replaceState(null,'','?host='+h+'&view='+view);}catch(ignore){}}
  // 功能区再次点击入口时在当前页面内切换，不重新加载，已填写的内容保留。
  window.OfficeToolsUi={run:run,showView:showView};PaneBridge.listen(showView);
  setView(new URLSearchParams(location.search).get('view')||'quick');refresh();setInterval(refresh,2500);
  if(OfficeTools.startLinkWatch)OfficeTools.startLinkWatch();$('loadHint').remove();
  // 实测 WPS 侧栏里只靠 DOMContentLoaded 时的翻译没有生效，这里在界面搭好后再明确翻一次。
  if(I18n)try{I18n.start();}catch(ignore){}
  if(OfficeTools.reportHost)OfficeTools.reportHost();
  // 切换语言后：重新生成由代码拼出的文字（卡片信息、当前文档标签等）。
  if(I18n)I18n.onChange(function(){cardsSignature='';refresh();renderCards();renderQuick();renderManage();});
})();

