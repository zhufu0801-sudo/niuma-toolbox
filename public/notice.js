// 完成提示（2026-11-24）：WPS 的 alert 只有黄色警告图标，成功时改用这个小窗口，显示绿色对勾。
// 文字通过地址里的 msg 传进来（中文原文），这里按界面语言翻译。
(function(){'use strict';
  var p=new URLSearchParams(location.search),msg=String(p.get('msg')||'').slice(0,4000),t=function(s){return window.OfficeI18n?OfficeI18n.t(s):s;};
  var box=document.getElementById('text'),ok=document.getElementById('ok');
  function render(){box.textContent=t(msg);ok.textContent=t('确定');}
  render();if(window.OfficeI18n&&OfficeI18n.onChange)OfficeI18n.onChange(function(){setTimeout(render,0);});
  function close(){try{if(window.Application&&Application.ActiveDialog&&Application.ActiveDialog.Close){Application.ActiveDialog.Close();return;}}catch(ignore){}try{window.close();}catch(ignore){}}
  ok.onclick=close;document.addEventListener('keydown',function(e){if(e.key==='Enter'||e.key==='Escape')close();});
  try{ok.focus();}catch(ignore){}
})();
