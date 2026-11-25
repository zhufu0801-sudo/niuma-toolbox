// 办公工具箱 · 界面语言（中文 / 日文）。
// 代码里的文字仍写中文；日文模式下按词典把界面文字、提示和对话框翻译成日文。
// 词典在 i18n-ja.js：键是代码里的中文原文（整句或拼接用的片段），值是日文。
(function(root){
  'use strict';
  var JA=root.OfficeI18nJa||{},RULES=root.OfficeI18nJaRules||[],KEY='officeTools.lang',MARK='⁣',listeners=[];
  var has=Object.prototype.hasOwnProperty,pref='auto',lang='zh',pattern=null;
  function esc(s){return s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');}
  function system(){var list=[];try{list=root.navigator.languages&&root.navigator.languages.length?root.navigator.languages:[root.navigator.language];}catch(ignore){}for(var i=0;i<list.length;i++){var l=String(list[i]||'').toLowerCase();if(l.indexOf('ja')===0)return 'ja';if(l.indexOf('zh')===0)return 'zh';}return 'zh';}
  var hinted=null;
  function resolve(p){return p==='ja'||p==='zh'?p:(hinted||system());}
  try{pref=root.localStorage.getItem(KEY)||'auto';}catch(ignore){}
  if(['auto','zh','ja'].indexOf(pref)<0)pref='auto';lang=resolve(pref);
  function fragments(){if(pattern===null){var keys=Object.keys(JA).filter(function(k){return k&&JA[k]!==undefined;}).sort(function(a,b){return b.length-a.length;});pattern=keys.length?new RegExp(keys.map(esc).join('|'),'g'):false;}return pattern;}
  // 用户自己的内容（卡片标题、文件名等）用 MARK 包起来，翻译时原样保留。
  function mark(text){return MARK+String(text==null?'':text)+MARK;}
  function piece(s){var re=fragments();if(re)s=s.replace(re,function(m){return JA[m];});return s.replace(/“/g,'「').replace(/”/g,'」').replace(/，/g,'、');}
  function translate(text){
    if(text===null||text===undefined)return text;
    // 词典里整句就有的（例如带示例路径的提示），直接用，不先把路径抽走。
    if(lang==='ja'){var whole=String(text),tm=whole.match(/^(\s*)([\s\S]*?)(\s*)$/);if(tm[2]&&has.call(JA,tm[2]))return tm[1]+JA[tm[2]]+tm[3];}
    var paths=[],s=String(text).replace(/(?:[A-Za-z]:[\\/]|\\\\|\/(?:Users|home|Volumes|tmp|var)\/)[^\r\n<>|"⁣]*/g,function(p){paths.push(p);return '\uE100'+(paths.length-1)+'\uE101';});
    // 只还原本次替换的占位符；多行消息逐行再翻译时，外层的占位符要留给外层还原。
    return translateCore(s).replace(/\uE100(\d+)\uE101/g,function(m,i){return Number(i)<paths.length?paths[Number(i)]:m;});
  }
  function translateCore(text){
    if(text===null||text===undefined)return text;var s=String(text);
    if(lang!=='ja')return s.split(MARK).join('');
    var m=s.match(/^(\s*)([\s\S]*?)(\s*)$/),lead=m[1],body=m[2],tail=m[3];
    if(!body)return s;var plain=body.split(MARK).join('');
    if(has.call(JA,plain))return lead+JA[plain]+tail;
    for(var i=0;i<RULES.length;i++){var r=RULES[i];if(r[0].test(plain)){r[0].lastIndex=0;return lead+plain.replace(r[0],r[1])+tail;}}
    if(!/[一-鿿]/.test(plain))return lead+plain+tail;
    var parts=body.split(MARK);for(var k=0;k<parts.length;k+=2)parts[k]=piece(parts[k]);
    return lead+parts.join('')+tail;
  }
  function extend(entries){Object.keys(entries).forEach(function(k){if(!has.call(JA,k))JA[k]=entries[k];});pattern=null;rev=null;}
  // 有的脚本把已经翻成日文的提示直接写进页面（t(...) 之后再赋值）。记原文时反查回中文，切回中文界面才能跟着变。
  // 2026-11-20 测试：设置页切到中文后「バージョン情報の URL が未設定です。」仍是日文。
  var rev=null;function source(v){if(typeof v!=='string')return v;var m=v.match(/^(\s*)([\s\S]*?)(\s*)$/);if(!m[2]||!/[぀-ヿ]/.test(m[2]))return v;if(!rev){rev={};Object.keys(JA).forEach(function(k){var j=JA[k];if(typeof j==='string'&&j&&!has.call(rev,j))rev[j]=k;});}return has.call(rev,m[2])?m[1]+rev[m[2]]+m[3]:v;}
  // ---- 页面文字：保存中文原文，切换语言时从原文重新翻译 ----
  var originals=typeof WeakMap==='function'?new WeakMap():null,written=typeof WeakMap==='function'?new WeakMap():null,ATTRS=['placeholder','title','aria-label','alt','label'];
  // 只认 body 里面元素上的 translate="no"：有的宿主会在 <html>/<body> 上加 translate="no"（防止浏览器自动翻译），不能因此整页不翻。
  function skip(node){var doc=root.document,top=doc&&doc.body;for(var el=node.nodeType===1?node:node.parentNode;el&&el.nodeType===1&&el!==top&&el!==(doc&&doc.documentElement);el=el.parentNode){var tag=el.tagName;if(tag==='SCRIPT'||tag==='STYLE'||tag==='TEXTAREA')return true;if(el.getAttribute&&el.getAttribute('translate')==='no')return true;}return false;}
  function textNode(node){if(!originals||skip(node))return;var value=node.nodeValue;if(written.get(node)!==value)originals.set(node,source(value));var next=translate(originals.get(node));if(next!==value){written.set(node,next);node.nodeValue=next;}else written.set(node,value);}
  // 输入框里的文字是用户写的，不翻译；但输入框自己的 placeholder、title 要翻。
  function element(el){if(!originals||(el.tagName==='TEXTAREA'?(el.parentNode&&el.parentNode.nodeType===1&&skip(el.parentNode)):skip(el)))return;var saved=originals.get(el)||{};ATTRS.forEach(function(name){if(!el.hasAttribute||!el.hasAttribute(name))return;var value=el.getAttribute(name),last=written.get(el)||{};if(last[name]!==value)saved[name]=source(value);var next=translate(saved[name]);last[name]=next;written.set(el,last);if(next!==value)el.setAttribute(name,next);});originals.set(el,saved);}
  function walk(rootNode){if(!rootNode)return;if(rootNode.nodeType===3){textNode(rootNode);return;}if(rootNode.nodeType!==1&&rootNode.nodeType!==9&&rootNode.nodeType!==11)return;if(rootNode.nodeType===1)element(rootNode);var doc=rootNode.ownerDocument||rootNode;if(!doc.createTreeWalker)return;var w=doc.createTreeWalker(rootNode,5,null),n;while((n=w.nextNode())){if(n.nodeType===3)textNode(n);else element(n);}}
  var observer=null,titleSource=null;
  function applyDocument(){var doc=root.document;if(!doc||!doc.documentElement)return;doc.documentElement.lang=lang==='ja'?'ja':'zh-CN';if(titleSource===null)titleSource=doc.title;doc.title=translate(titleSource);if(doc.body)walk(doc.body);}
  function watch(){var doc=root.document;if(observer||!doc||!doc.body||typeof MutationObserver!=='function')return;observer=new MutationObserver(function(list){list.forEach(function(m){if(m.type==='characterData')textNode(m.target);else if(m.type==='attributes')element(m.target);else m.addedNodes.forEach(walk);});});observer.observe(doc.body,{subtree:true,childList:true,characterData:true,attributes:true,attributeFilter:ATTRS});}
  function start(){applyDocument();watch();}
  // ---- 语言选择：本机缓存一份，正式保存在设置里 ----
  // 各窗口（功能区主页面、侧栏、设置对话框）是分开的网页：选好的语言另存一份到 WPS 的 PluginStorage 和本机缓存，各窗口每 2 秒对一次。
  // 只接受本页打开之后别的窗口发来的切换：打开前留下的旧值不算数，以设置里保存的为准（实测旧值会把“跟随系统”又改回日文）。
  var SHARED='officeTools.lang.shared',applied=Date.now();
  function plugin(){try{var a=root.OfficeTools&&root.OfficeTools.app?root.OfficeTools.app():root.Application||(root.wps&&root.wps.Application);return a&&a.PluginStorage&&typeof a.PluginStorage.getItem==='function'?a.PluginStorage:null;}catch(ignore){return null;}}
  function share(p){var item=JSON.stringify({p:p,at:Date.now()}),ps=plugin();applied=Date.now();try{if(ps)ps.setItem(SHARED,item);}catch(ignore){}try{root.localStorage.setItem(SHARED,item);}catch(ignore){}}
  function shared(){var best=null;[function(){return plugin().getItem(SHARED);},function(){return root.localStorage.getItem(SHARED);}].forEach(function(get){try{var v=JSON.parse(get());if(v&&typeof v.p==='string'&&(!best||v.at>best.at))best=v;}catch(ignore){}});return best;}
  // mode：save 保存到设置并通知其他窗口；sync 来自设置或其他窗口；preview 只改当前页面（设置页里试看）。
  function setPreference(p,mode){if(['auto','zh','ja'].indexOf(p)<0)p='auto';var before=lang,oldPref=pref;pref=p;lang=resolve(p);if(mode!=='preview')try{root.localStorage.setItem(KEY,p);}catch(ignore){}if(mode==='save'){share(p);if(root.WorkbenchSettings)root.WorkbenchSettings.update({ui:{language:p}}).catch(function(){});}var prefChanged=before!==lang||mode==='save'||oldPref!==pref;applyDocument();if(prefChanged)listeners.forEach(function(fn){try{fn(lang,pref);}catch(ignore){}});return lang;}
  function pollShared(){var v=shared();if(v&&v.at>applied){applied=v.at;if(v.p!==pref)setPreference(v.p,'sync');}}
  if(typeof setInterval==='function')setInterval(function(){try{pollShared();}catch(ignore){}},2000);
  // WPS 自己的界面语言比浏览器语言更准：日文版 WPS 在“跟随系统”时显示日文。
  function hint(l){if(l!=='ja'&&l!=='zh'||hinted===l)return;hinted=l;if(pref==='auto'){var before=lang;lang=resolve(pref);applyDocument();if(before!==lang)listeners.forEach(function(fn){try{fn(lang,pref);}catch(ignore){}});}}
  function bindSettings(){var s=root.WorkbenchSettings;if(!s||bindSettings.done)return;bindSettings.done=true;function sync(data){var p=data&&data.settings&&data.settings.ui&&data.settings.ui.language;if(p&&p!==pref)setPreference(p,'sync');}s.subscribe(sync);sync(s.snapshot());}
  root.OfficeI18n={debug:function(){var nav='';try{nav=(root.navigator.languages||[root.navigator.language]).join(',');}catch(ignore){}var snap='',ls='';try{var d=root.WorkbenchSettings&&root.WorkbenchSettings.snapshot();snap=d?String(d.settings&&d.settings.ui&&d.settings.ui.language)+'@'+d.revision:'none';}catch(e){snap='err';}try{ls=root.localStorage.getItem(KEY);}catch(e){ls='err';}return 'wrap='+wrapNote.join(',')+';hint='+hinted+';sys='+system()+';nav='+nav+';bound='+!!bindSettings.done+';snap='+snap+';ls='+ls+';applied='+applied;},t:translate,mark:mark,lang:function(){return lang;},preference:function(){return pref;},system:system,set:function(p){return setPreference(p,'save');},use:function(p){return setPreference(p,'preview');},share:share,hint:hint,onChange:function(fn){listeners.push(fn);},apply:walk,start:start,bindSettings:bindSettings};
  root.OfficeI18n.extend=extend;
  // 浏览器里的提示框也跟随语言。
  // 实测 WPS 的功能区主页面：改写 window.alert 会报错，报错会让后面的初始化全部跳过，所以每一步都单独兜住。
  var wrapNote=[],natives={};['alert','confirm'].forEach(function(name){try{var original=root[name];if(typeof original!=='function'){wrapNote.push(name+':none');return;}natives[name]=original;var fn=function(message){return original.call(root,translate(message));};try{root[name]=fn;}catch(e){wrapNote.push(name+':set:'+e.message);}if(root[name]!==fn)try{Object.defineProperty(root,name,{value:fn,configurable:true,writable:true});}catch(e){wrapNote.push(name+':def:'+e.message);}if(root[name]!==fn)wrapNote.push(name+':kept');}catch(e){wrapNote.push(name+':'+e.message);}});
  // 改写不了时（WPS 功能区主页面就是这样），代码里统一调用 OfficeI18n.alert / OfficeI18n.confirm，也能显示翻译后的提示。
  ['alert','confirm'].forEach(function(name){root.OfficeI18n[name]=function(message){var f=natives[name]||root[name];return f.call(root,translate(message));};});
  // WPS 的 alert 一律显示黄色警告图标，换不了；成功的提示在前面加“✓”，和“未完成：”区分开（2026-11-20 测试 P3）。
  root.OfficeI18n.done=function(message){var f=natives.alert||root.alert;return f.call(root,'✓ '+translate(message));};
  // 实测 WPS 侧栏：页面脚本跑完时文档还处于 loading，DOMContentLoaded 之后的翻译没有生效过；
  // 所以 DOMContentLoaded、load 和稍后的计时器都各翻一次（翻译可以重复执行，不会叠加）。
  function kick(){try{start();}catch(ignore){}bindSettings();}
  if(root.document){if(root.document.readyState==='loading')root.document.addEventListener('DOMContentLoaded',kick);else{kick();}if(root.addEventListener)root.addEventListener('load',kick);setTimeout(kick,0);setTimeout(kick,1500);}
  setTimeout(bindSettings,0);
})(window);
