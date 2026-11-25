// 设置页“AI 服务”：选择服务商、填写 Key、获取模型列表、测试连接。Key 只发给本机服务保存，页面读不回明文。
(function(){
 'use strict';var $=function(id){return document.getElementById(id);},token='',providers=[],config={},privacy={};
 // 页面文字由 i18n.js 自动翻成日文，这里不要再翻一次（否则日文会被二次替换）。
 function t(x){return x;}
 function msg(text,error){var b=$('aiMessage');b.textContent=t(text||'');b.className=error?'error':text?'ok':'';}
 async function api(url,body,retry){if(!token)token=(await fetch('/api/session').then(function(r){return r.json();})).token;var r=await fetch(url,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify(body||{})});if(r.status===403&&!retry){token='';return api(url,body,true);}var d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(d.error||'操作未完成，请重试。');return d;}
 function ja(){var i=window.OfficeI18n;return !!(i&&i.lang&&i.lang()==='ja');}
 function provider(){return providers.find(function(p){return p.id===$('aiProvider').value;})||{};}
 function fill(){var sel=$('aiProvider');sel.textContent='';var off=document.createElement('option');off.value='none';off.textContent=ja()?'（AI を使わない）':'（不使用 AI）';sel.append(off);// 日文界面把日本客户常用的服务放前面。
  var JA_FIRST=['openai','claude','gemini'],list=providers.slice();if(ja())list.sort(function(a,b){var x=JA_FIRST.indexOf(a.id),y=JA_FIRST.indexOf(b.id);return (x<0?99:x)-(y<0?99:y);});
  list.filter(function(p){return !p.noKey||$('aiAdvanced').checked||p.id===config.provider;}).forEach(function(p){var o=document.createElement('option');o.value=p.id;o.textContent=ja()?p.nameJa:p.name;sel.append(o);});
   // 没有选服务商（或点过“不使用 AI”）时，显示“不使用 AI”，避免看起来还连着原来的服务。
   sel.value=config.provider||'none';$('aiBase').value=config.baseUrl||provider().baseUrl||'';$('aiModel').value=config.model||provider().model||'';
   $('aiPrivacy').value=(config.privacy&&config.privacy.mode)||'mask';
   var pl=privacy.plugins||[],bad=privacy.failed||[];$('aiPrivacyInfo').textContent=t('发送前会把身份证号、手机号、邮箱、银行卡号、个人编号换成占位符，收到回答后再换回原文。')+(pl.length?t('已加载隐私插件：')+pl.map(function(x){return x.name+(x.version?' '+x.version:'');}).join('、')+t('。'):'')+(bad.length?t('有插件没有加载成功，AI 功能已暂停：')+bad[0].file:'');
   $('aiKeyInfo').textContent=config.hasKey?(config.keyProtection==='dpapi'?t('API Key 已用 Windows 加密保存，只有当前 Windows 用户能解开。'):config.keyProtection==='keychain'?t('API Key 已存进 Mac 的钥匙串，只有当前 Mac 用户能读取。'):t('API Key 以明文保存在本机。')+(config.keyWarning?'（'+t(config.keyWarning)+'）':'')):'';
   $('aiKey').value='';$('aiKey').placeholder=t(config.hasKey?'已保存：':'粘贴 API Key')+(config.hasKey?config.keyHint+t('（不改就留空）'):'');link();}
 function link(){var p=provider(),a=$('aiKeyLink');a.hidden=!p.keyUrl;a.href=p.keyUrl||'#';a.textContent=t(p.noKey?'下载 Ollama':'去官网申请 Key');$('aiKey').disabled=!!p.noKey;}
 $('aiAdvanced').onchange=function(){var saved={base:$('aiBase').value,model:$('aiModel').value,key:$('aiKey').value,privacy:$('aiPrivacy').value};config.provider=$('aiProvider').value;fill();$('aiBase').value=saved.base;$('aiModel').value=saved.model;$('aiKey').value=saved.key;$('aiPrivacy').value=saved.privacy;};
 $('aiProvider').onchange=function(){var p=provider();$('aiBase').value=p.baseUrl||'';$('aiModel').value=p.model||'';$('aiModels').textContent='';link();msg(p.model||$('aiProvider').value==='none'?'':'这家的模型名经常更新：填好 Key 并保存后，点“获取模型列表”选一个。');};
 function body(){var b={provider:$('aiProvider').value,baseUrl:$('aiBase').value,model:$('aiModel').value,privacy:{mode:$('aiPrivacy').value}};if($('aiKey').value.trim())b.apiKey=$('aiKey').value.trim();return b;}
 async function save(silent){var d=await api('/api/ai/config',body());config=d.config;fill();if(!silent)msg('AI 设置已保存。'+(config.ready?'可以点“测试连接”确认。':'还差模型或 Key，填好后再保存。'));return d;}
 $('aiSave').onclick=function(){save().catch(function(e){msg(e.message,true);});};
 $('aiPrivacy').onchange=function(){save(true).then(function(){msg('个人信息保护设置已保存。');}).catch(function(e){msg(e.message,true);});};
 $('aiClear').onclick=function(){api('/api/ai/config',{provider:$('aiProvider').value,baseUrl:$('aiBase').value,model:$('aiModel').value,privacy:{mode:$('aiPrivacy').value},clearKey:true}).then(function(d){config=d.config;fill();msg('已清除保存的 API Key。');}).catch(function(e){msg(e.message,true);});};
 $('aiOff').onclick=function(){api('/api/ai/config',{provider:'none'}).then(function(d){config=d.config||{};fill();msg('已停用 AI：工具箱里的 AI 功能不会再连接任何服务。需要时重新选服务商并保存即可。');}).catch(function(e){msg(e.message,true);});};
 $('aiList').onclick=async function(){try{msg('正在读取模型列表…');if($('aiKey').value.trim()||$('aiBase').value!==config.baseUrl||$('aiProvider').value!==config.provider)await save(true);var d=await api('/api/ai/models');var dl=$('aiModels');dl.textContent='';d.models.forEach(function(m){var o=document.createElement('option');o.value=m;dl.append(o);});msg(d.models.length?'模型列表已更新（'+d.models.length+' 个），点“模型”输入框即可选择。':'服务商没有返回模型列表，请按官网文档填写模型名。');}catch(e){msg(e.message,true);}};
 $('aiTest').onclick=async function(){try{msg('正在测试连接…');await save(true);var d=await api('/api/ai/test');if(d.ok)msg('连接成功：'+d.model+'，用时 '+(d.ms/1000).toFixed(1)+' 秒。'+(d.thinking?'这是思考型模型，回答会慢一些。':''));else{if(d.models&&d.models.length){var dl=$('aiModels');dl.textContent='';d.models.forEach(function(m){var o=document.createElement('option');o.value=m;dl.append(o);});}msg(d.error+(d.models&&d.models.length?' 已列出可用模型，请在“模型”框里重新选择。':''),true);}}catch(e){msg(e.message,true);}};
 fetch('/api/ai/config').then(function(r){return r.json();}).then(function(d){providers=d.providers||[];config=d.config||{};privacy=d.privacy||{};fill();if(!config.provider)msg('目前没有使用 AI。要用的话，选服务商、填好后点保存。');}).catch(function(){msg('无法读取 AI 设置，请确认办公工具箱仍在运行。',true);});
 if(window.OfficeI18n&&window.OfficeI18n.onChange)window.OfficeI18n.onChange(function(){if(providers.length)fill();});
})();
