(function(){
  'use strict';var pane,busy=false,SWITCH_WAIT=1500,currentView='quick';
  var RIBBON=window.OfficeRibbonText||{};
  var statusOwner='',statusLanguage='';
  function documentKey(){try{var a=OfficeTools.app(),h=OfficeTools.host(),d=h==='et'?a.ActiveWorkbook:h==='wpp'?a.ActivePresentation:a.ActiveDocument;return h+'|'+String(d.FullName||d.Name);}catch(ignore){return '';}}
  function clearStaleStatus(){if(statusOwner&&(statusOwner!==documentKey()||statusLanguage!==lang())){try{OfficeTools.app().StatusBar=false;}catch(ignore){}statusOwner='';}}
  if(typeof setInterval==='function')setInterval(clearStaleStatus,500);
  if(window.OfficeI18n&&OfficeI18n.onChange)OfficeI18n.onChange(clearStaleStatus);
  // WPS 功能区主页面里 window.alert 是只读的，翻译不能靠改写它，提示统一走 say()。
  function say(message){var i=window.OfficeI18n;return i&&typeof i.alert==='function'?i.alert(message):alert(message);}
  // 成功提示：WPS 的 alert 只有黄色警告图标（2026-11-20 客户体验 P3），改用带绿色对勾的小窗口；打不开时再用 alert。
  function done(message){var text=String(message||'');try{var a=OfficeTools.app();if(a&&typeof a.ShowDialog==='function'&&text.length<=3000){var lines=text.split('\n').length,h=Math.min(460,170+lines*22);a.ShowDialog(location.origin+'/notice.html?msg='+encodeURIComponent(text),t('办公工具箱'),460,h,true,true,0,'',10000,true,false,true);return;}}catch(ignore){}var i=window.OfficeI18n;return i&&typeof i.done==='function'?i.done(text):say(text);}
  function t(text){return window.OfficeI18n?OfficeI18n.t(text):text;}
  // TaskPane has no writable title API. Recreate only on an actual language change.
  if(window.OfficeI18n&&OfficeI18n.onChange)OfficeI18n.onChange(function(){if(!pane)return;try{var visible=pane.Visible;pane.Delete();pane=null;if(visible)create(currentView);}catch(ignore){pane=null;}});
  // 「常用工具」组里按程序放的常用按钮：id 前面加 cw_（文字）/ ce_（表格）/ cp_（演示），文字、图标、动作都用原来那个按钮的（2026-11-25）。
  var COMMON_HOST={cw:'wps',ce:'et',cp:'wpp'};
  function baseId(id){return String(id||'').replace(/^c[wep]_/,'');}
  function lang(){try{return window.OfficeI18n?OfficeI18n.lang():'zh';}catch(ignore){return 'zh';}}
  function fav(id){return /^fav[1-6]$/.test(String(id||''))&&window.QuickFavorites?QuickFavorites.slot(id):null;}
  window.GetLabel=function(control){var f=fav(control&&control.Id);if(f)return t(f.label);var e=RIBBON[control&&control.Id==='sheetCompactGroup'?'sheetGroup':baseId(control&&control.Id)];return e&&e.label?t(e.label):'';};
  window.GetScreentip=function(control){var f=fav(control&&control.Id);if(f)return t(f.desc)+(f.action?'':t('（打开操作指引）'));var e=RIBBON[baseId(control&&control.Id)];return e&&e.tip?t(e.tip):'';};
  // 侧栏里开关功能后，功能区按钮跟着显示或隐藏：定时比较开关的修改时间，有变化就刷新功能区。
  var ribbon=null,prefsAt=0;
  // 功能区图标只能返回不带参数的相对路径（icons/*.svg），WPS 才会到本机服务取图。
  // 实测（2026-11-08，WPS 12.1.0.29690，诊断里的 recentFiles 记录为证）：加 ?r= 参数、完整网址、本地路径、file 地址、data 地址，WPS 都不取图，按钮显示“?”。
  // getImage 只在第一次显示选项卡时调用；WPS 先于本机服务启动时图标也会是“?”，重启 WPS 即可。
  window.OnAddinLoad=function(ui){ribbon=ui||null;if(OfficeTools.startLinkWatch)OfficeTools.startLinkWatch();try{prefsAt=OfficeTools.prefsStamp?OfficeTools.prefsStamp():0;}catch(ignore){}if(window.OfficeI18n&&ribbon&&typeof ribbon.Invalidate==='function')OfficeI18n.onChange(function(){try{ribbon.Invalidate();}catch(ignore){}});if(ribbon&&typeof ribbon.Invalidate==='function'&&typeof setInterval==='function')setInterval(function(){try{var at=OfficeTools.prefsStamp();if(at!==prefsAt){prefsAt=at;ribbon.Invalidate();}}catch(ignore){}},2000);try{OfficeTools.reportHost();}catch(ignore){}return true;};window.OnAddInLoad=window.OnAddinLoad;
  window.GetImage=function(control){var id=control&&control.Id;if(/^fav[1-6]$/.test(String(id||''))){var f=fav(id);return f&&f.action?ToolIcons.forAction(f.action):f?'icons/cards.svg':'icons/star.svg';}return ToolIcons.forAction(baseId(id));};
  window.GetVisible=function(control){var h=OfficeTools.host(),id=control.Id;if(id==='sheetCompactGroup')return h==='et';var cm=/^(c[wep])_(.+)$/.exec(id);if(cm){if(COMMON_HOST[cm[1]]!==h)return false;if(cm[2]==='officialFormatCn')return lang()!=='ja';if(cm[2]==='businessFormatJa')return lang()==='ja';try{return !(OfficeTools.isHidden&&OfficeTools.isHidden(cm[2]));}catch(ignore){return true;}}if(id==='studioOpen')return h!=='wpp'&&!(OfficeTools.isHidden&&OfficeTools.isHidden(id));if(id==='printKeepRows')return h==='wps';if(id==='printSheetSetup')return h==='et';if(id==='wordGroup'||id==='docGroup')return h==='wps';if(/^fav[1-6]$/.test(id))return !!fav(id);if(id==='favHint'){try{return !(window.QuickFavorites&&QuickFavorites.favorites(h).length);}catch(ignore){return true;}}if(h==='et'&&['spotlight','sheetCheck','formulaOpen'].indexOf(id)>=0)return false;if(id==='sheetGroup'||id==='sheetPlusGroup')return h==='et';if(id==='batchReplace'||id==='batchWatermark'||id==='splitByHeading')return h==='wps';if(id==='slideGroup'||id==='slideOutGroup')return h==='wpp';if(id==='proofGroup')return h==='wps';if(id==='sheetDataGroup')return h==='et';if(id==='mailMergeOpen')return h==='wps'&&!(OfficeTools.isHidden&&OfficeTools.isHidden(id));var only={dateSummary:['et'],unifyDates:['et','wps'],approvalBox:['et','wps']}[id];if(only&&only.indexOf(h)<0)return false;try{return !(OfficeTools.isHidden&&OfficeTools.isHidden(id));}catch(ignore){return true;}};
  function paneKey(){return 'officeTools.paneId.'+OfficeTools.host();}
  function paneUrl(view){return location.origin+'/tools.html?host='+OfficeTools.host()+'&view='+PaneBridge.view(view);}
  function remember(p){try{OfficeTools.app().PluginStorage.setItem(paneKey(),String(p.ID));}catch(ignore){}}
  function recall(){try{var a=OfficeTools.app(),id=a.PluginStorage.getItem(paneKey());return id?a.GetTaskPane(Number(id)||id):null;}catch(ignore){return null;}}
  function create(view){var p=OfficeTools.app().CreateTaskPane(paneUrl(view),t('办公工具箱'));if(!p)throw Error('当前 WPS 未能创建工具侧栏。');try{p.Width=380;}catch(ignore){}p.Visible=true;pane=p;remember(p);}
  // 侧栏已存在则只显示出来；已被关闭或失效则返回 false，由调用方重新创建。
  function show(){if(!pane)pane=recall();if(!pane)return false;try{pane.Visible=true;return true;}catch(ignore){pane=null;return false;}}
  function open(view){
    view=PaneBridge.view(view);currentView=view;
    if(!show()){create(view);return Promise.resolve();}
    return PaneBridge.request(view,SWITCH_WAIT).then(function(switched){if(!switched)throw Error('工具侧栏已显示，但页面切换暂未响应。已保留填写内容，请点击侧栏顶部的“常用 / 制作 / 文件”标签。');});
  }
  function openStudio(){
    var result=StudioLaunch.open({app:OfficeTools.app(),origin:location.origin,open:window.open.bind(window)});
    if(!result.ok)say(result.message+'\n手动地址：'+result.url);
  }
  // WPS 启动后会一直保留这里加载的脚本。更新文件后若没重启 WPS，按钮会悄悄跑旧代码，所以先核对版本。
  var checkedAt=0,lastOk=true;
  function fresh(){
    if(!OfficeTools.version||typeof fetch!=='function')return Promise.resolve(true);
    // 只缓存“版本一致”的结果；不一致时每次点击都提示，不能静默失败。
    if(lastOk&&Date.now()-checkedAt<60000)return Promise.resolve(true);
    return fetch('version.json',{cache:'no-store'}).then(function(r){return r.ok?r.json():null;}).then(function(info){
      checkedAt=Date.now();lastOk=!info||!info.version||info.version===OfficeTools.version;
      if(!lastOk)say('办公工具箱已更新到 '+info.version+'，但 WPS 里还是旧版本 '+OfficeTools.version+'。\n请保存文档，关闭所有 WPS 窗口后重新打开。');
      return lastOk;
    },function(){return true;});
  }
  window.OpenQuickNavigator=function(id){try{OfficeTools.app().PluginStorage.setItem('officeTools.quickRequest.'+OfficeTools.host(),JSON.stringify({id:id,at:Date.now()}));open('quick').catch(function(e){say(e.message);});}catch(e){say(e.message);}};
  // “文档整理”的结果里有处理了多少处、哪些需要手动检查，弹出来而不只是写在状态栏。
  function docfix(id){try{return OfficeTools.list(OfficeTools.host()).some(function(x){return x.id===id&&(x.group==='docfix'||x.group==='sheetplus'||x.group==='print');});}catch(ignore){return false;}}
  // 需要填写内容的工具（前后缀、改名、分隔符）：打开侧栏技巧导航里的这一项，填好后在侧栏执行。
  // 第二批新功能里先要选一下怎么做的（检查方式、编号样式、导出方式等），点功能区按钮时先打开侧栏里的这一项。
  var FORM_FIRST={postalFill:1,jaStyle:1,proofCheck:1,captionNumbers:1,fixListNumbering:1,wanUnit:1,idCheck:1,mergeWorkbooks:1,codesBatch:1,slidesToImages:1,handoutPdf:1};
  function needsForm(id){try{var x=OfficeTools.list(OfficeTools.host()).find(function(t){return t.id===id;});return !!(x&&x.fields&&x.fields.some(function(f){return f.required!==false;}));}catch(ignore){return false;}}
  window.OnAction=function(control){
    var id=baseId(control.Id);
    if(/^fav[1-6]$/.test(id)){var f=fav(id);if(!f)return true;if(!f.action){window.OpenQuickNavigator(f.id);return true;}id=f.action;}
    if(id==='formulaOpen'){window.OpenQuickNavigator('formulaAssistant');return true;}
    var aiPanel={postalOpen:'jumpPostal',macroCheck:'jumpMacroCheck',dataMask:'jumpMask',ocrOpen:'ocrImage',pdfToWordOpen:'pdfToWord',imagesExtractOpen:'extractImages',imagesCompressOpen:'compressImages',historyOpen:'docHistory',formsOpen:'jumpForms',mailMergeOpen:'jumpMailMerge',pdfToolsOpen:'jumpPdfTools',workdayCalc:'workdayCalc',aiChatOpen:'aiChat',aiSummary:'aiSummary',aiContinue:'aiContinue',aiExpand:'aiExpand',aiShorten:'aiShorten',aiBullets:'aiBullets',aiFormal:'aiFormal',aiTranslate:'aiTranslate'}[id];if(aiPanel){window.OpenQuickNavigator(aiPanel);return true;}
    if(needsForm(id)||FORM_FIRST[id]){window.OpenQuickNavigator(id);return true;}
    if(id==='favHint'){try{open('quick').then(null,function(e){say(e.message);});}catch(e){say(e.message);}return true;}
    if(id==='printFilesOpen'){open('files').catch(function(e){say(e.message);});return true;}
    if(id==='pasteBarOpen'){open('files').catch(function(e){say(e.message);});return true;}
    if(id==='settingsOpen'){try{OfficeTools.app().ShowDialog(location.origin+'/settings.html',t('办公工具箱设置'),760,760,false);}catch(e){window.open(location.origin+'/settings.html','_blank');}return true;}
    if(id==='studioOpen'){openStudio();return true;}
    if(id==='toolsOpen'||id==='insertTable'||id==='makeSlides'||id==='openCards'){
      try{open(id==='openCards'?'cards':id==='makeSlides'||id==='insertTable'?'create':'quick').then(null,function(e){say(e.message);});}catch(e){say(e.message);}
      return true;
    }
    if(busy)return true;busy=true;var actionOwner=documentKey();
    fresh().then(function(ok){if(!ok){busy=false;return;}
    OfficeTools.runFromUi(id).then(function(message){try{if(actionOwner===documentKey()){OfficeTools.app().StatusBar=t(message);statusOwner=actionOwner;statusLanguage=lang();}}catch(ignore){}if(id==='capture'||id==='pasteBarCollect'||id==='saveCard'||id==='exportPdf'||id==='insertTransfer'||id==='insertLatest'||id==='backupFile'||id==='saveLinkedModule'||id==='updateLinkedModules'||docfix(id)||/^(没有|当前选区没有|选区内没有)/.test(String(message)))(/^(没有|当前选区没有|选区内没有)/.test(String(message))?say:done)(message);},function(e){var pre=/^(请|选区|当前|没有|两边|表头|这个|已有一个|按 Ctrl|“)/.test(String(e.message));say('未完成：'+e.message+(pre?'':OfficeTools.host()==='et'?'\n如表格已有部分更改，请检查内容；表格中由工具写入的内容可能无法用撤销恢复。':'\n如文档已有部分更改，可使用 WPS 自带的撤销按钮。'));}).then(function(){busy=false;});
    });
    return true;
  };
})();
