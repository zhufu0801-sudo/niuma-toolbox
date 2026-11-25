(function(root){
  'use strict';
  // 打开 PPT 版式工作室。WPS 内用 ShowDialog；没有 WPS 对象时视为浏览器预览，改用新标签页。
  var NOTE='浏览器版可以预览并下载新的 PPTX，但不会修改 WPS 里当前打开的文档。';
  function open(env){
    env=env||{};
    var app=env.app,url=env.origin+'/studio.html';
    if(!app){
      var tab=null;
      try{tab=env.open(url,'_blank');}catch(ignore){}
      if(tab){try{tab.opener=null;}catch(ignore){}return {ok:true,mode:'browser',url:url,message:'已在浏览器新标签页打开版式工作室。'+NOTE};}
      return {ok:false,mode:'browser',url:url,message:'浏览器拦截了新窗口。请允许本页弹出窗口后重试，或点击下方“在浏览器中打开”。'};
    }
    if(typeof app.ShowDialog!=='function')return {ok:false,mode:'wps',kind:'unsupported',url:url,message:'当前 WPS 版本不支持内置制作窗口。可点击下方“在浏览器中打开”制作新的 PPTX。'+NOTE};
    // 窗口大小按屏幕算：高分屏（缩放 150% 等）上固定的 1040×760 会显得很小，要来回滚动（2026-11-20 小白测试）。
    var size=dialogSize(env.screen||root.screen,env.ratio||root.devicePixelRatio);
    try{app.ShowDialog(url,'PPT 版式工作室',size[0],size[1],false);}
    catch(e){return {ok:false,mode:'wps',kind:'error',url:url,message:'制作窗口没有打开：'+((e&&e.message)||'原因未知')+'。请再试一次；仍无反应时可点击下方“在浏览器中打开”。'+NOTE};}
    return {ok:true,mode:'wps',url:url,message:'已请求打开制作窗口。如果没有看到窗口，请再点一次，或点击下方“在浏览器中打开”。'};
  }
  function dialogSize(scr,ratio){var r=Number(ratio)||1,w=Math.round((scr&&scr.availWidth||1280)*r*0.85),h=Math.round((scr&&scr.availHeight||800)*r*0.88);return [Math.max(1040,Math.min(w,2400)),Math.max(760,Math.min(h,1500))];}
  root.StudioLaunch={open:open,note:NOTE,dialogSize:dialogSize};
})(window);
