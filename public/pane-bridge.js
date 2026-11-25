(function(root){
  'use strict';
  // 功能区脚本与已打开的工具侧栏同源但各自独立。这里只传递“要显示哪一页”这一个值，
  // 侧栏收到后在原页面内切换，不重新加载，所以已填写的大纲不会丢。
  var VIEWS=['quick','create','cards','files'],STALE_MS=10000;
  function view(value){return VIEWS.indexOf(value)<0?'quick':value;}
  function host(){try{return root.OfficeTools?root.OfficeTools.host():'wps';}catch(ignore){return 'wps';}}
  function name(kind){return 'officeTools.pane.'+kind+'.'+host();}
  function pluginStore(){try{var a=root.OfficeTools&&root.OfficeTools.app();return a&&a.PluginStorage&&typeof a.PluginStorage.getItem==='function'?a.PluginStorage:null;}catch(ignore){return null;}}
  function write(key,value){var text=JSON.stringify(value),p=pluginStore();try{if(p)p.setItem(key,text);}catch(ignore){}try{root.localStorage.setItem(key,text);}catch(ignore){}}
  function read(key){
    var best=null;
    [function(){return pluginStore().getItem(key);},function(){return root.localStorage.getItem(key);}].forEach(function(get){try{var item=JSON.parse(get());if(item&&typeof item.id==='string'&&(!best||item.at>best.at))best=item;}catch(ignore){}});
    return best;
  }
  function channel(){try{return typeof root.BroadcastChannel==='function'?new root.BroadcastChannel(name('channel')):null;}catch(ignore){return null;}}
  function newId(){return Date.now().toString(36)+Math.random().toString(36).slice(2,8);}
  // 功能区一侧：请已打开的侧栏切换到指定页面。收到回执返回 true；超时（侧栏未响应）返回 false。
  function request(target,timeout){
    var message={type:'request',id:newId(),view:view(target),at:Date.now()},wire=channel();
    return new Promise(function(resolve){
      var done=false,timer,poller;
      function finish(ok){if(done)return;done=true;clearTimeout(timer);clearInterval(poller);if(wire){wire.onmessage=null;wire.close();}resolve(ok);}
      function check(){var ack=read(name('ack'));if(ack&&ack.id===message.id)finish(true);}
      if(wire)wire.onmessage=function(event){if(event.data&&event.data.type==='ack'&&event.data.id===message.id)finish(true);};
      poller=setInterval(check,100);timer=setTimeout(function(){finish(false);},timeout||1500);
      write(name('request'),message);
      if(wire)wire.postMessage(message);
    });
  }
  // 侧栏一侧：收到切换请求就调用 onView，成功后回执；onView 抛错则不回执，让功能区改用重建。
  function listen(onView){
    var current=read(name('request')),seen=current&&current.id,wire=channel(),poller;
    function handle(message){
      if(!message||message.type!=='request'||message.id===seen)return;
      seen=message.id;
      if(Date.now()-message.at>STALE_MS)return;
      try{onView(view(message.view));}catch(ignore){return;}
      var ack={type:'ack',id:message.id,at:Date.now()};write(name('ack'),ack);if(wire)wire.postMessage(ack);
    }
    function pull(){handle(read(name('request')));}
    if(wire)wire.onmessage=function(event){handle(event.data);};
    try{root.addEventListener('storage',pull);}catch(ignore){}
    poller=setInterval(pull,400);
    return function stop(){clearInterval(poller);try{root.removeEventListener('storage',pull);}catch(ignore){}if(wire){wire.onmessage=null;wire.close();}};
  }
  root.PaneBridge={view:view,request:request,listen:listen};
})(window);
