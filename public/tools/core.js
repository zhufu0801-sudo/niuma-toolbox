// 办公工具箱 · 核心模块：共用的读取与检查函数、功能登记表、执行入口和功能开关。
// 其他模块（word / sheet / slides / transfer）通过 OfficeTools.register 登记自己的功能。
(function(root){
  'use strict';
  // 宿主名称拼进提示里；日文界面直接给日文产品名，并用标记保护，避免被当成普通词再翻译一次。
  var NAMES={wps:['文字','Writer'],et:['表格','Spreadsheets'],wpp:['演示','Presentation']},labels={};
  function i18n(){return root.OfficeI18n||null;}
  function mark(text){var t=i18n();return t?t.mark(text):String(text);}
  Object.keys(NAMES).forEach(function(h){Object.defineProperty(labels,h,{enumerable:true,get:function(){var t=i18n();return t&&t.lang()==='ja'?mark(NAMES[h][1]):NAMES[h][0];}});});
  var VERSION='0.16.0',RESTART='如果刚更新过办公工具箱，请保存文档，关闭所有 WPS 窗口后重新打开。';
  function app(){return root.Application||(root.wps&&root.wps.Application)||root.wps;}
  function host(){var m=location.pathname.match(/\/addon\/(wps|et|wpp)\//);return m?m[1]:(new URLSearchParams(location.search).get('host')||'wps');}
  function active(a,h){var d=h==='wps'?a.ActiveDocument:h==='et'?a.ActiveWorkbook:a.ActivePresentation;if(!d)throw Error('请先打开一份'+labels[h]+'文档。');if(d.ReadOnly)throw Error('当前文档为只读，请另存为可编辑副本后操作。');return d;}
  function identity(d){return String(d.FullName||d.Name);}
  function wordSelection(a){var s=a.Selection;if(!s||!String(s.Text||'').replace(/[\r\n\u0007\s]/g,''))throw Error('请先在正文里选中要处理的文字。');if(String(s.Text).length>50000)throw Error('一次最多处理 5 万字，请缩小选区。');return s;}
  function cells(a){var r=a.Selection;if(!r||!r.Rows||!r.Columns)throw Error('请先选择单元格区域。');if(r.Areas&&r.Areas.Count>1)throw Error('请先选择一片连续的单元格区域。');if(r.Rows.Count>1000||r.Columns.Count>50)throw Error('选区超过 1000 行或 50 列，请缩小范围。');if(a.ActiveSheet.ProtectContents)throw Error('工作表已受保护，请在可编辑副本中操作。');return r;}
  function rgb(r,g,b){return r+g*256+b*65536;}
  function tableInput(options){var text=String(options.text||'').replace(/\r/g,'').replace(/^\n+|\n+$/g,'');if(!text.trim())throw Error('请先填写表格内容。');var rows=text.split('\n').map(function(line){return line.split('\t');});var cols=Math.max.apply(null,rows.map(function(row){return row.length;}));if(rows.length>20||cols>10)throw Error('快速插入最多 20 行、10 列，请拆分内容。');if(text.length>10000)throw Error('内容过长，请缩小范围。');rows.forEach(function(row){while(row.length<cols)row.push('');});return rows;}
  function localFetch(url,options){var control=typeof AbortController==='function'?new AbortController():null,timer=control&&setTimeout(function(){control.abort();},15000);if(control)options.signal=control.signal;return fetch(url,options).then(function(r){if(control)clearTimeout(timer);return r;},function(e){if(control)clearTimeout(timer);throw Error(e&&e.name==='AbortError'?'本地服务长时间没有响应，请稍后重试；如仍无反应，请重新打开办公工具箱。':'无法连接本地服务，请确认办公工具箱仍在运行后重试。');});}
  async function api(method,body){var session=await localFetch('/api/session',{}).then(function(r){if(!r.ok)throw Error('本地服务未连接。');return r.json();});var r=await localFetch('/api/local-transfer',{method:method,headers:{'Content-Type':'application/json','X-Coach-Token':session.token},body:method==='POST'?JSON.stringify(body):undefined});var data=await r.json().catch(function(){return {};});if(!r.ok)throw Error(data.error||'内容传递失败，请重试。');return data;}
  function selectedEnd(a){var r=a.Selection.Range.Duplicate;r.Collapse(0);return r;}
  function replaceOptions(options){var find=String(options.findText||''),replacement=String(options.replaceText||'');if(!find)throw Error('请填写要查找的文字。');if(find.length>200||replacement.length>1000)throw Error('查找文字最多 200 字，替换文字最多 1000 字。');if(/[\r\n\u0007]/.test(find+replacement))throw Error('快速替换暂不处理段落或表格分隔符，请只输入普通文字。');return {find:find,replacement:replacement};}
  function position(a,h,d){var result=identity(d);if(h==='et'){var selection=cells(a);result+='|'+a.ActiveSheet.Name+'|'+String(typeof selection.Address==='function'?selection.Address():selection.Address);}if(h==='wps')result+='|'+a.Selection.Range.Start+'|'+a.Selection.Range.End;return result;}
  // ---- 功能登记 ----
  // 每个功能：id、名称、说明、图标、分组，以及各宿主（wps 文字 / et 表格 / wpp 演示）的执行函数。
  // 同一 id 可以在不同模块里分别登记不同宿主的实现，会合并到一起。
  var actions={},order=[];
  function register(def){
    var cur=actions[def.id];
    var first=!cur;
    if(first){cur=actions[def.id]={id:def.id,hosts:{}};order.push(def.id);}
    ['label','desc','icon','group','fields','keywords'].forEach(function(k){if(def[k]!==undefined)cur[k]=def[k];});
    // ownConfirm：这些宿主里工具自己会弹详细确认，不再先弹一次通用确认。
    if(def.ownConfirm)cur.ownConfirm=(cur.ownConfirm||[]).concat(def.ownConfirm);
    if(first)cur.panel=def.panel!==false;
    // panel 只作用于这次登记的宿主：例如表格里的“匹配周围格式”只是提示用格式刷，不放进常用列表。
    Object.keys(def.hosts||{}).forEach(function(h){cur.hosts[h]={fn:def.hosts[h],undo:def.undo!==false,label:def.labels&&def.labels[h],desc:def.descs&&def.descs[h],panel:first?undefined:def.panel};});
    return cur;
  }
  function describe(id,h){var x=actions[id];if(!x)return null;var e=x.hosts[h]||{};return {id:id,label:e.label||x.label,desc:e.desc||x.desc,icon:x.icon,group:x.group,panel:e.panel!==undefined?e.panel!==false:x.panel,available:!!x.hosts[h],fields:x.fields||null,keywords:x.keywords||'',ownConfirm:x.ownConfirm||null};}
  function list(h){return order.filter(function(id){return actions[id].hosts[h];}).map(function(id){return describe(id,h);});}
  // ---- 功能开关（用户自己选择显示哪些功能）----
  var HIDDEN_KEY='officeTools.hidden';
  function store(){try{var a=app();return a&&a.PluginStorage&&typeof a.PluginStorage.getItem==='function'?a.PluginStorage:null;}catch(ignore){return null;}}
  function legacyPrefs(){var best={ids:[],at:0};[function(){return store().getItem(HIDDEN_KEY);},function(){return root.localStorage.getItem(HIDDEN_KEY);}].forEach(function(get){try{var v=JSON.parse(get());if(v&&Array.isArray(v.ids)&&v.at>=best.at)best=v;}catch(ignore){}});return best;}
  var bridge=root.WorkbenchSettings;
  if(bridge)bridge.legacy(legacyPrefs);
  function readPrefs(){var state=bridge&&bridge.snapshot();return state?{ids:state.settings.hiddenIds,at:state.revision}:legacyPrefs();}
  // 日文界面且仍是默认中文字体时，改用日文 Windows 自带的字体，避免日文排版落到宋体。
  var JA_FONTS={'宋体':'Yu Mincho','微软雅黑':'Yu Gothic','黑体':'Yu Gothic'};
  function localFont(name){var t=i18n();return t&&t.lang()==='ja'&&JA_FONTS[name]?JA_FONTS[name]:name;}
  function settings(){var state=bridge&&bridge.snapshot(),s=state?JSON.parse(JSON.stringify(state.settings)):{text:{font:'宋体',size:12,lineSpacing:1.5},table:{headerBackground:'#DDEBF7',headerColor:'#1F3864'},cards:{limit:40}};if(s.text)s.text.font=localFont(s.text.font);return s;}
  function ready(){return bridge?bridge.ready():Promise.resolve();}
  function hidden(){return readPrefs().ids;}
  function prefsStamp(){return readPrefs().at;}
  function setHidden(ids){ids=ids.filter(function(id){return typeof id==='string';});function legacy(){var v=JSON.stringify({ids:ids,at:Date.now()}),p=store();try{if(p)p.setItem(HIDDEN_KEY,v);}catch(ignore){}try{root.localStorage.setItem(HIDDEN_KEY,v);}catch(ignore){}}if(bridge)return bridge.update({hiddenIds:ids}).then(legacy);legacy();return Promise.resolve();}
  function isHidden(id){return hidden().indexOf(id)>=0;}
  // ---- 宿主版本与接口检测 ----
  // 正版 WPS 有个人版、专业版/企业版、海外版、日本版等，新旧版本的接口不完全一样。
  // 启动时读出版本号、界面语言和关键接口是否可用，上报给本机服务（启动页显示），功能按可用的接口选择做法。
  function probe(obj,name){try{return !!(obj&&obj[name]);}catch(ignore){return false;}}
  var LCID={1041:'ja',2052:'zh',1028:'zh',3076:'zh',4100:'zh'};
  function environment(){var a=app(),h=host(),info={host:h,version:'',build:'',name:'',language:'',caps:{}};if(!a)return info;
    try{info.version=String(a.Version||'');}catch(ignore){}try{info.build=String(a.Build||'');}catch(ignore){}try{info.name=String(a.Name||'');}catch(ignore){}
    try{var id=a.LanguageSettings&&a.LanguageSettings.LanguageID(2);info.language=LCID[id]||(id?String(id):'');}catch(ignore){}
    info.caps={taskPane:probe(a,'CreateTaskPane'),dialog:probe(a,'ShowDialog'),storage:!!store(),fileDialog:probe(a,'FileDialog'),undo:h==='wps'?probe(a,'UndoRecord'):null};
    // 日文版 WPS 且语言设为“跟随系统”时，界面直接用日文。
    var t=i18n();if(t&&t.hint&&(info.language==='ja'||info.language==='zh'))t.hint(info.language);
    return info;}
  
  function reportHost(){if(typeof fetch!=='function')return Promise.resolve(false);var info=environment(),i=root.OfficeI18n;info.page=pageKind();info.ui=i?{lang:i.lang(),pref:i.preference(),probe:(i.debug?i.debug()+'|':'')+diag()}:null;info.errors=errors.slice();return localFetch('/api/session',{}).then(function(r){return r.json();}).then(function(s){return localFetch('/api/host-info',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(info)});}).then(function(r){return !!(r&&r.ok);},function(){return false;});}
  // 诊断：本页的语言状态和最近的脚本错误一起报给本机服务（只存在服务内存里，排查用）。
  var errors=[];if(root.addEventListener){root.addEventListener('error',function(e){errors.push(String(e&&e.message||e).slice(0,200));if(errors.length>5)errors.shift();});root.addEventListener('unhandledrejection',function(e){errors.push(String(e&&e.reason&&e.reason.message||e&&e.reason||'').slice(0,200));if(errors.length>5)errors.shift();});}
  function diag(){try{var d=root.document,h=d&&d.querySelector('h2,h1');return [root.OfficeI18n.t('常用工具'),d.documentElement.getAttribute('translate'),d.documentElement.lang,h?h.textContent:'',typeof WeakMap,typeof MutationObserver,d.readyState].join('|').slice(0,500);}catch(e){return 'err:'+e.message;}}
  // 确认框：WPS 功能区主页面里改写不了 window.confirm，所以统一走这里，日文模式下也能显示日文。
  // 表格里撤销不了：文件已存到电脑且开着自动备份时，告诉客户已经自动备份；否则提醒先保存。
  function backupNote(){try{var s=bridge&&bridge.snapshot(),auto=!(s&&s.settings&&s.settings.backups&&s.settings.backups.auto===false),d=active(app(),'et'),p=String(d&&d.FullName||'');if(auto&&/^([A-Za-z]:[\\/]|\/|\\\\)/.test(p))return '\nWPS 表格中的改动不能用 Ctrl+Z 恢复；执行前会自动备份磁盘上的这个文件（在设置的备份文件夹里）。';}catch(ignore){}return '\nWPS 表格中的改动不能用 Ctrl+Z 恢复。这个文件还没有保存过，无法自动备份，建议先保存。';}
  function ask(message){if(host()==='et'&&!/撤销|Ctrl\+Z/.test(message))message+=backupNote();var i=root.OfficeI18n;return i&&typeof i.confirm==='function'?i.confirm(message):root.confirm(message);}
  function pageKind(){try{var p=String(root.location.pathname||'');return /tools\.html$/.test(p)?'pane':/settings\.html$/.test(p)?'settings':'main';}catch(ignore){return 'main';}}
  if(root.OfficeI18n&&root.OfficeI18n.onChange)root.OfficeI18n.onChange(function(){try{reportHost();}catch(ignore){}});
  if(root.addEventListener)root.addEventListener('load',function(){setTimeout(function(){try{reportHost();}catch(ignore){}},3000);setTimeout(function(){try{reportHost();}catch(ignore){}},12000);});
  // 自动备份：文档已经存到电脑上时，先请本机服务复制一份磁盘上的文件（修改还没保存，所以不用等复制完成）。
  function autoBackup(d){try{var s=bridge&&bridge.snapshot();if(s&&s.settings&&s.settings.backups&&s.settings.backups.auto===false)return;var p=String(d.FullName||'');if(!/^([A-Za-z]:[\\/]|\/|\\\\)/.test(p)||typeof fetch!=='function')return;
    localFetch('/api/session',{}).then(function(r){return r.json();}).then(function(x){return localFetch('/api/tools/auto-backup',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':x.token},body:JSON.stringify({path:p})});}).catch(function(){});}catch(ignore){}}
  // ---- 执行入口 ----
  async function run(id,options){
    await ready();options=options||{};var a=app(),h=options.host||host();if(!a)throw Error('请在 WPS 顶部“办公工具箱”中使用此功能，浏览器页面无法修改文档。');var d=active(a,h),record=null;
    var action=actions[id],entry=action&&action.hosts[h];
    if(!entry)throw Error('当前'+labels[h]+'窗口不支持这个功能。'+RESTART);
    try{
      if(h==='wps'&&entry.undo){try{record=a.UndoRecord;record.StartCustomRecord('办公工具箱');}catch(ignore){record=null;}}
      autoBackup(d);
      return await entry.fn({a:a,d:d,h:h,options:options});
    }finally{if(record)try{record.EndCustomRecord();}catch(ignore){}}
  }
  async function runFromUi(id,options){var h=host(),tool=describe(id,h);if(['capture','pasteBarCollect','saveCard','backupFile','exportPdf','printPreflight','sheetCheck','rangeToImage','exportSheets','batchPdf','saveLinkedModule'].indexOf(id)<0&&!(tool&&tool.ownConfirm&&tool.ownConfirm.indexOf(h)>=0)){var note=h==='wpp'?'\n部分演示操作无法撤销，建议先保存副本。':'';if(!ask('将执行“'+(tool&&tool.label||id)+'”，修改当前文档或选区。是否继续？'+note))return '已取消。';}return run(id,options);}
  root.OfficeTools={run:run,runFromUi:runFromUi,host:host,app:app,tableInput:tableInput,version:VERSION,register:register,list:list,describe:describe,hidden:hidden,setHidden:setHidden,isHidden:isHidden,prefsStamp:prefsStamp,ready:ready,settings:settings,environment:environment,reportHost:reportHost,
    lib:{labels:labels,mark:mark,ask:ask,localFont:localFont,RESTART:RESTART,active:active,identity:identity,wordSelection:wordSelection,cells:cells,rgb:rgb,tableInput:tableInput,localFetch:localFetch,api:api,selectedEnd:selectedEnd,replaceOptions:replaceOptions,position:position}};
})(window);
