(function(){
 'use strict';var $=function(id){return document.getElementById(id);},token='',ready=false;
 async function api(url,method,retry){if(!token)token=(await fetch('/api/session').then(function(r){return r.json();})).token;var r=await fetch(url,{method:method||'GET',headers:{'X-Coach-Token':token,'Content-Type':'application/json'},body:method==='POST'?'{}':undefined});if(r.status===403&&!retry){token='';return api(url,method,true);}var data=await r.json();if(!r.ok)throw Error(data.error||'本地服务暂时无法连接。');return data;}
 function row(id,value,ok,sub){$('row-'+id).dataset.state=ok?'ok':'warn';$('val-'+id).textContent=value;var icon=$('row-'+id).querySelector('use');if(icon)icon.setAttribute('href',ok?'#i-ok':'#i-warn');if(sub){$('sub-'+id).hidden=false;$('sub-'+id).textContent=sub;}}
 // WPS 版本与兼容情况：区分个人版/按电脑安装（专业版、企业版常见），中国版/海外版/日本版。
 var REGION={cn:'中国版',intl:'海外版',jp:'日本版'},SCOPE={user:'个人版（当前用户安装）',machine:'按电脑安装（专业版、企业版常见）'},HOST={wps:'文字',et:'表格',wpp:'演示'},CAPS={taskPane:'侧栏',dialog:'对话框',storage:'插件存储',fileDialog:'选择文件夹'};
 function hostName(h){var I=window.OfficeI18n;return I&&I.lang()==='ja'?I.mark({wps:'Writer',et:'Spreadsheets',wpp:'Presentation'}[h]):HOST[h];}
 function describeWps(env,found){var hosts=env.hosts||{},loaded=Object.keys(hosts),install=(env.installs||[])[0],notes=[],ok=found;
  var report=loaded.length?hosts[loaded[0]]:{},full=function(v){return /^\d+\.\d+\.\d+\.\d+$/.test(String(v||''));};
  // Application.Version 只给出“12.0”，完整版本号在 Build 里；都没有时用安装目录里的版本。
  var version=full(report.build)?report.build:install&&full(install.version)?install.version:report.version||install&&install.version;
  var label=version?'WPS '+version:(found?'已找到 WPS':'未找到 WPS');var kind=install?[REGION[install.region],SCOPE[install.scope]].filter(Boolean).join(' · '):'';
  if(loaded.length)notes.push('工具箱已在 WPS '+loaded.map(function(h){return hostName(h);}).join('、')+' 中加载');else if(found)notes.push('打开 WPS 后，这里会显示工具箱是否已加载');
  if(install&&install.major&&install.major<11){notes.push('这个 WPS 版本较旧，可能无法加载工具箱，建议升级到 WPS 2019（11）或更新版本');ok=false;}
  if(install&&install.jsApiPlugin===false){notes.push('这台电脑的 WPS 关闭了加载项功能，请让管理员在 oem.ini 的 [Support] 下设置 JsApiPlugin=true');ok=false;}
  var reg=env.registration||{};if(reg.found){var missing=Object.keys(HOST).filter(function(h){return !reg.hosts[h];});if(missing.length){notes.push('工具箱还没有登记到 WPS '+missing.map(function(h){return hostName(h);}).join('、'));ok=false;}}
  loaded.forEach(function(h){var caps=hosts[h].caps||{},lack=Object.keys(CAPS).filter(function(k){return caps[k]===false;});if(lack.length)notes.push('WPS '+hostName(h)+' 缺少'+lack.map(function(k){return CAPS[k];}).join('、')+'接口，相关功能会改用备用方式');});
  row('wps',label,ok,[kind].concat(notes).filter(Boolean).join('\n'));}
 // Microsoft Office：只显示检测结果；Office 版插件由安装程序（Start.cmd）选择后登记。
 function describeOffice(o){var el=document.getElementById('row-office');if(!el)return;if(!o||!o.installed){row('office','未检测到 Microsoft Office',true,'需要时可重新运行 Start.cmd 选择安装到 Office。');return;}
  row('office',o.name+(o.version?' '+o.version:''),!!o.supported,o.supported?'Office 版为试用版：在 Excel / Word 顶部的“办公工具箱”标签打开侧栏。没有这个标签时，重新运行 Start.cmd 并勾选 Microsoft Office。':'版本太旧，Office 版需要 2016 及以后（建议 Microsoft 365）。');}
 async function check(){ready=false;$('startBtn').disabled=true;$('recheckBtn').disabled=true;try{var data=await Promise.all([fetch('/api/health',{cache:'no-store'}).then(function(r){if(!r.ok)throw Error('服务未连接');return r.json();}),fetch('/version.json',{cache:'no-store'}).then(function(r){return r.json();}),api('/api/startup')]);var health=data[0],version=data[1],state=data[2],match=health.app==='wps-ai-coach'&&health.version===version.version;row('svc',match?'已连接':'请重启服务',match);$('version').textContent='版本 '+health.version;document.querySelectorAll('.ver-text').forEach(function(el){el.textContent=health.version;});$('versionBanner').hidden=match;$('versionDetail').textContent='当前服务 '+health.version+'，页面 '+version.version+'。请重新启动办公工具箱。';row('wps',state.wpsFound?'已找到 WPS':'未找到 WPS',state.wpsFound);var office=null;try{var env=await api('/api/environment');describeWps(env,state.wpsFound);office=env.office;}catch(ignore){}describeOffice(office);row('cards',state.cardsCount+' 张内容卡片',true);row('backup','备份保存位置',true,state.backupDir);ready=match&&state.wpsFound;var officeOnly=match&&!state.wpsFound&&office&&office.supported;$('startHint').textContent=ready?'打开 WPS 后，点击顶部“办公工具箱”。':officeOnly?'请打开 Excel 或 Word，点击顶部“办公工具箱”标签。':'检查上方状态，处理后点击“重新检查”。';}catch(e){row('svc','无法连接',false);$('startHint').textContent=e.message+' 请重新打开桌面上的办公工具箱。';}finally{$('startBtn').disabled=!ready;$('recheckBtn').disabled=false;}}
 $('startBtn').onclick=async function(){$('startBtn').disabled=true;$('launchMsg').hidden=false;try{await api('/api/launch-wps','POST');$('launchText').textContent='已请求打开 WPS，请在顶部选择“办公工具箱”。';}catch(e){$('launchText').textContent='未能打开：'+e.message;}finally{$('startBtn').disabled=!ready;}};
 // 语言切换：保存到设置里，侧栏和功能区也会跟着变。
 function markLang(){var p=window.OfficeI18n?OfficeI18n.preference():'auto';document.querySelectorAll('[data-lang]').forEach(function(b){b.setAttribute('aria-pressed',String(b.dataset.lang===p));});}
 document.querySelectorAll('[data-lang]').forEach(function(b){b.onclick=function(){if(window.OfficeI18n){OfficeI18n.set(b.dataset.lang);markLang();}};});if(window.OfficeI18n)OfficeI18n.onChange(function(){markLang();
  // 状态里的应用名称（文字/Writer 等）是按当时的语言拼出来的，换语言后重新检查一遍。
  check();});markLang();
 $('recheckBtn').onclick=check;document.querySelectorAll('[data-info]').forEach(function(button){button.onclick=function(){var panel=$('info-'+button.dataset.info);if(panel){panel.hidden=!panel.hidden;button.setAttribute('aria-expanded',String(!panel.hidden));}};});check();
})();
