// 办公工具箱 · Microsoft Office 快捷键。
// 键位固定为 Ctrl+Alt+Shift+1 … 0（写在 shortcuts.json 里，Office 只在安装时读一次）。每个键执行哪个工具可以在侧栏里改，保存在本侧栏（每台电脑各自一份）。
// 需要 Office 支持“共享运行时”和自定义快捷键：Excel 2102、Word 2408、PowerPoint 2601 及以后的 Microsoft 365。安装程序检测到够新的版本才会打开这个功能。
// 尚未在真实 Microsoft Office 中验证。
(function(root){
  'use strict';
  var SLOTS=['1','2','3','4','5','6','7','8','9','0'];
  var DEFAULTS={
    word:{'1':'deleteBlankPages','2':'officialFormatCn','3':'tableBeautify','4':'pictureFix','5':'fixPageJump','6':'boldInlineHeads','7':'removeWatermark','8':'alignOptions','9':'addCjkSpace','0':'openPane'},
    excel:{'1':'spotlight','2':'bandRows','3':'textToNumber','4':'deleteBlankRows','5':'compareRanges','6':'trimCells','7':'printTitles','8':'sheetCheck','9':'backupFile','0':'openPane'},
    powerpoint:{'1':'formatSlide','2':'duplicateSlide','9':'backupFile','0':'openPane'}
  };
  // 需要填写内容的工具（查找替换、按大纲生成）不能用快捷键直接执行。
  var NEEDS_FORM={replaceText:1,makeSlides:1,insertCard:1,addPrefixSuffix:1,renameSheets:1};
  function store(){var W=root.WorkbenchSettings,s=W&&W.snapshot();return s&&s.settings&&s.settings.officeKeys||{};}
  function slots(host){var saved=store()[host]||{},def=DEFAULTS[host]||{},out={};SLOTS.forEach(function(k){out[k]=Object.prototype.hasOwnProperty.call(saved,k)?saved[k]:(def[k]||'');});return out;}
  function keyOf(slot){return 'Ctrl+Alt+Shift+'+slot;}
  function choices(host,list){return [{id:'',label:'（不使用）'},{id:'openPane',label:'打开工具箱侧栏'}].concat((list||[]).filter(function(x){return !NEEDS_FORM[x.id]&&!x.form&&!(x.fields&&x.fields.some(function(f){return f.required!==false;}));}).map(function(x){return {id:x.id,label:x.label};}));}
  function set(host,slot,id){if(SLOTS.indexOf(String(slot))<0)return Promise.reject(Error('快捷键编号不正确。'));var all=store(),cur=Object.assign({},slots(host));cur[slot]=id||'';var patch={};Object.keys(all).forEach(function(h){patch[h]=all[h];});patch[host]=cur;return root.WorkbenchSettings.update({officeKeys:patch});}
  function reset(host){var all=store(),patch={};Object.keys(all).forEach(function(h){if(h!==host)patch[h]=all[h];});patch[host]=Object.assign({},DEFAULTS[host]||{});return root.WorkbenchSettings.update({officeKeys:patch});}
  // run(id) 由侧栏提供：执行工具；返回 Promise。Office 要求回调返回 Promise 或 undefined。
  function associate(office,host,run){var a=office&&office.actions;if(!a||typeof a.associate!=='function')return false;
    SLOTS.forEach(function(k){try{a.associate('toolbox.k'+k,function(){var id=slots(host)[k];if(!id)return Promise.resolve();return Promise.resolve(run(id)).then(function(){return undefined;},function(e){return String(e&&e.message||e);});});}catch(e){}});return true;}
  root.OfficeKeys={SLOTS:SLOTS,DEFAULTS:DEFAULTS,slots:slots,keyOf:keyOf,choices:choices,set:set,reset:reset,associate:associate};
})(typeof window!=='undefined'?window:globalThis);
