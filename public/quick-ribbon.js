(function(root){
 'use strict';function esc(s){return String(s).replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c];});}
 // 能直接执行的条目：已有工具，或带 run 的方法卡（“一键完成”）。其余条目（需要选参数的快捷工具、操作指引）打开技巧导航。
 function action(x){return x?(x.run||(x.kind==='existing'?x.id:'')):'';}
 function toolkit(){var W=root.WorkbenchSettings||(typeof WorkbenchSettings!=='undefined'?WorkbenchSettings:null),s=W&&W.snapshot();return s&&s.settings.toolkit||{};}
 var MAX=6;
 // 功能区“我的收藏”最多 6 个，按收藏顺序。
 function favorites(host){var ids=toolkit().favorites||[];if(!ids.length)return [];var all=QuickCatalog.list(host),out=[];ids.forEach(function(id){var x=all.find(function(c){return c.id===id;});if(x&&out.length<MAX)out.push({id:x.id,label:x.label,desc:x.desc,action:action(x)});});return out;}
 function slot(id){var m=/^fav([1-6])$/.exec(String(id||''));return m?favorites(OfficeTools.host())[Number(m[1])-1]||null:null;}
 root.QuickFavorites={favorites:favorites,slot:slot,action:action,MAX:MAX};
 root.GetQuickMenu=function(){var host=OfficeTools.host(),bindings=toolkit().bindings||[],fav=toolkit().favorites||[];var t=function(x){return root.OfficeI18n?root.OfficeI18n.t(x):x;};
  var list=QuickCatalog.list(host).filter(function(x){return x.menu!==false||x.run||fav.indexOf(x.id)>=0;});list.sort(function(a,b){return Number(fav.indexOf(b.id)>=0)-Number(fav.indexOf(a.id)>=0);});
  var rows=list.map(function(x){var b=bindings.find(function(b){return b.host===host&&b.id===x.id;});return '<button id="qh_'+esc(x.id)+'" label="'+esc((fav.indexOf(x.id)>=0?'★ ':'')+t(x.label))+'" screentip="'+esc(t(x.desc))+'" onAction="OnQuickMenuAction"'+(b&&b.tip?' keytip="'+esc(b.tip)+'"':'')+'/>';});return '<menu xmlns="http://schemas.microsoft.com/office/2006/01/customui">'+rows.join('')+'</menu>';};
 root.OnQuickMenuAction=function(control){var id=String(control.Id||'').replace(/^qh_/,''),x=QuickCatalog.list(OfficeTools.host()).find(function(x){return x.id===id;});if(!x)return true;var run=action(x);if(run&&typeof root.OnAction==='function')root.OnAction({Id:run});else root.OpenQuickNavigator(id);return true;};
})(window);
