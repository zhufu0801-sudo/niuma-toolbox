(function(root){'use strict';
// 首次使用引导。2026-11-20 体验测试：原来是系统默认的灰按钮、小字，最后一页还有“跳过”；
// 现在按侧栏的样式画，最后一页问一句要不要开“新手模式”（只显示最常用的工具），选完就记住。
var KEY='officeToolbox.welcome.v1';
var PAGES=[
 ['欢迎使用办公工具箱','在 WPS 顶部找到“办公工具箱”标签。在 Office 的“加载项”中打开侧栏。'],
 ['从当前文档开始','打开侧栏，选择常用工具。修改前请核对选区，重要文件先备份。'],
 ['不会用就直接搜','在“你想做什么？”里用自己的话搜，比如“格式乱”“删重复”“转成 PDF”。'],
 ['文件与图片','图片识别、PDF 转 Word、压缩图片会在本机处理，结果另存为副本。'],
 ['AI 可选','不用 AI 也能使用整理工具。需要 AI 时，在设置中填写自己的 API Key；调用可能产生费用。'],
 ['要不要先只显示常用工具？','刚开始用的话，可以先只显示最常用的十来个工具，界面更清爽。以后在“设置 → 显示哪些功能”里随时可以改。']
];
function t(s){var i=root.OfficeI18n;return i?i.t(s):s;}
function button(text,primary){var b=document.createElement('button');b.type='button';b.textContent=t(text);b.style.cssText='display:block;width:100%;margin:10px 0 0;padding:10px 12px;font-size:15px;line-height:1.4;border-radius:8px;cursor:pointer;'+(primary?'background:#2466ab;color:#fff;border:1px solid #2466ab;font-weight:600':'background:#fff;color:#2466ab;border:1px solid #858f9c');return b;}
function setBeginner(on){try{if(root.WorkbenchSettings)root.WorkbenchSettings.update({ui:{beginner:!!on}}).catch(function(){});}catch(e){}}
function show(){var old=document.getElementById('toolboxWelcome');if(old)old.remove();var n=0,last=PAGES.length-1,box=document.createElement('section');box.id='toolboxWelcome';box.setAttribute('role','dialog');box.setAttribute('aria-label',t('首次使用引导'));
 box.style.cssText='position:fixed;inset:12px;z-index:10000;background:#fff;color:#16263b;padding:24px 20px;border:2px solid #2466ab;border-radius:14px;overflow:auto;box-shadow:0 8px 40px #0004;font-size:15px;line-height:1.7';
 function close(){try{localStorage.setItem(KEY,'1');}catch(e){}box.remove();}
 function draw(){box.textContent='';var h=document.createElement('h2'),p=document.createElement('p'),count=document.createElement('p');h.textContent=t(PAGES[n][0]);h.style.cssText='font-size:20px;margin:0 0 12px';p.textContent=t(PAGES[n][1]);p.style.cssText='margin:0 0 8px';count.textContent=(n+1)+' / '+PAGES.length;count.style.cssText='color:#5b6675;font-size:13px;margin:0 0 6px';box.append(h,p,count);
  if(n===last){var easy=button('只显示常用工具（推荐新手）',true),all=button('显示全部工具',false);easy.onclick=function(){setBeginner(true);close();};all.onclick=function(){setBeginner(false);close();};box.append(easy,all);easy.focus();return;}
  var next=button('下一页',true),skip=button('跳过引导',false);next.onclick=function(){n++;draw();};skip.onclick=close;box.append(next,skip);next.focus();}
 document.body.append(box);draw();}
root.ToolboxWelcome={show:show,pages:PAGES};
if(document.getElementById('showWelcome'))document.getElementById('showWelcome').onclick=show;else{try{if(!localStorage.getItem(KEY))show();}catch(e){}}
})(window);
