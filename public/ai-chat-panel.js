// AI 对话：在侧栏里像聊天一样问 AI（操作方法、公式、起草文字）。可以附带当前选中的内容；每条回答可以复制或插入文档。
// 使用“设置 → AI 服务”里接入的 AI，发送前经过隐私保护。对话只保存在本侧栏，关闭侧栏即清空。
(function(root){
 'use strict';
 var T=root.OfficeTools,P=root.AiTextPanel;if(!T||!P)return;
 var MAX_TURNS=30,history=[],busy=false;
 var section=document.createElement('section');section.className='tool-card ai-chat-panel';section.dataset.panel='quick';section.id='aiChatPanel';
 section.innerHTML='<details><summary><strong>AI 对话</strong><span>问操作方法、写公式、起草文字 · 回答可以插入文档</span></summary>'
  +'<p class="input-hint" id="acNote"></p>'
  +'<div id="acLog" class="ac-log" aria-live="polite"></div>'
  +'<label for="acInput">问题</label><textarea id="acInput" rows="3" maxlength="4000" placeholder="例如：怎样让表格每页都打印表头？（Ctrl+Enter 发送）"></textarea>'
  +'<label class="ac-check"><input type="checkbox" id="acContext"> 附带当前选中的内容</label>'
  +'<div class="button-pair"><button type="button" id="acSend" class="primary">发送</button><button type="button" id="acClear">清空对话</button></div><button type="button" id="acCancel" hidden>取消</button>'
  +'<div id="acStatus" role="status" aria-live="polite"></div></details>';
 var $=function(id){return section.querySelector('#ac'+id);};
 function place(){var anchor=document.getElementById('aiTextPanel');if(anchor&&anchor.parentNode){anchor.parentNode.insertBefore(section,anchor);return true;}return false;}
 if(!place())document.addEventListener('DOMContentLoaded',place);
 function lang(){return root.OfficeI18n&&root.OfficeI18n.lang&&root.OfficeI18n.lang()==='ja'?'ja':'zh';}
 // 界面文字由 i18n 自动翻译；AI 的回答和用户的问题标了 translate="no"，不会被改动。
 function tr(s){return s;}
 function status(s,err){$('Status').textContent=s?tr(s):'';$('Status').className='fa-status'+(err?' error':'');}
 fetch('/api/ai/config').then(function(r){return r.json();}).then(function(d){var c=d.config||{},p=(d.providers||[]).find(function(x){return x.id===c.provider;});
   if(!c.ready){$('Note').textContent=tr('还没有接入 AI：请到“设置 → AI 服务”填写服务商和 API Key。');return;}
   var mode=(d.privacy&&d.privacy.mode)||'mask',name=p?(lang()==='ja'&&p.nameJa?p.nameJa:p.name):c.provider;
   $('Note').textContent='将使用 '+name+'。'+(p&&p.region==='local'?'内容只在这台电脑上处理。':({mask:'问题和附带的内容会发送给该服务商；身份证号、手机号等会先自动隐藏。',remind:'问题和附带的内容会发送给该服务商；含个人信息时会先问你。',off:'问题和附带的内容会发送给该服务商；隐私保护已关闭。'}[mode]));}).catch(function(){});
 function copy(text){if(navigator.clipboard&&navigator.clipboard.writeText)return navigator.clipboard.writeText(text).catch(fallback);return Promise.resolve(fallback());
  function fallback(){var ta=document.createElement('textarea');ta.value=text;ta.style.position='fixed';ta.style.opacity='0';document.body.appendChild(ta);ta.select();var ok=false;try{ok=document.execCommand('copy');}catch(e){}ta.remove();if(!ok)throw Error('复制失败，请选中文字后按 Ctrl+C。');}}
 async function insert(text){var host=P.host();if(!host.insert)throw Error('这个程序暂时不支持插入，请复制后粘贴。');var r=await host.insert(text,false);if(r&&r.confirm){if(!(await T.lib.ask(r.confirm)))return '已取消。';r=await host.insert(text,true);}return r;}
 function render(){var log=$('Log');log.textContent='';history.forEach(function(m){var box=document.createElement('div');box.className='ac-msg '+(m.role==='user'?'ac-user':'ac-ai');var who=document.createElement('span');who.className='ac-who';who.textContent=tr(m.role==='user'?'你':'AI');var body=document.createElement('div');body.className='ac-text';body.setAttribute('translate','no');body.textContent=m.content;box.append(who);if(m.withContext){var tag=document.createElement('span');tag.className='ac-tag';tag.textContent=tr('附带了选中的内容');box.append(tag);}box.append(body);
   if(m.role==='assistant'){var bar=document.createElement('div');bar.className='ac-actions';[['复制',function(){copy(m.content).then(function(){status('已复制。');},function(e){status(e.message,true);});}],['插入文档',function(){insert(m.content).then(function(r){status(r);},function(e){status(e.message,true);});}]].forEach(function(x){var b=document.createElement('button');b.type='button';b.textContent=tr(x[0]);b.onclick=x[1];bar.append(b);});box.append(bar);}
   log.append(box);});log.hidden=!history.length;$('Clear').disabled=!history.length||busy;$('Send').disabled=busy;if(history.length)log.scrollTop=log.scrollHeight;}
 async function context(){var src=await P.host().read();if(src.kind==='cells')return (src.address?src.address+'\n':'')+src.items.map(function(x){return x.text;}).join('\n');return src.text;}
 async function send(){if(busy)return;var q=$('Input').value.trim();if(!q){status('请先输入问题。',true);return;}
  var ctx='';if($('Context').checked){try{ctx=await context();}catch(e){status(e.message,true);return;}if(ctx.length>8000){status('附带的选中内容太长了（超过 8000 字），请少选一些。',true);return;}}
  busy=true;history.push({role:'user',content:q,withContext:!!ctx});$('Input').value='';render();$('Cancel').hidden=false;var stop=P.waiting?P.waiting(function(t){status(t);},'AI 正在回答，'):function(){};if(!P.waiting)status('AI 正在回答，通常需要 5～60 秒…');
  try{var d=await P.api('/api/ai/chat',{messages:history.map(function(m){return {role:m.role,content:m.content};}),context:ctx,lang:lang()});
   history.push({role:'assistant',content:d.text});while(history.length>MAX_TURNS)history.shift();while(history.length&&history[0].role!=='user')history.shift();
   status(tr('已回答（'+d.model+'）。')+(d.privacy&&d.privacy.masked?tr('发送前隐藏了 '+d.privacy.masked+' 处个人信息，回答里已还原。'):''));}
  catch(e){history.pop();$('Input').value=q;if(/还没有设置 AI/.test(e.message)){status('还没接入 AI，所以这里没法回答。',true);suggestLocal(q);}else status(e.message,true);}
  finally{stop();$('Cancel').hidden=true;busy=false;render();}}
 // 没接 AI 时：按问题推荐工具箱里现成的工具，再给一个“去设置 AI”的按钮（新手不知道 API Key 是什么）。
 function suggestLocal(q){var old=section.querySelector('.ac-suggest');if(old)old.remove();var box=document.createElement('div');box.className='ac-suggest';var list=root.QuickPanel&&root.QuickPanel.suggest?root.QuickPanel.suggest(q,3):[];
  var head=document.createElement('p');head.textContent=tr(list.length?'可以先试试这些现成的工具：':'没找到现成的工具，换个说法再搜搜看，或者接入 AI 后再问。');box.append(head);
  list.forEach(function(x){var b=document.createElement('button');b.type='button';b.className='link-button';b.textContent=tr(x.label);b.title=tr(x.desc||'');b.onclick=function(){root.QuickPanel.open(x.id);};box.append(b);});
  var go=document.createElement('button');go.type='button';go.className='link-button';go.textContent=tr('去设置 AI（需要自己的 AI 账号）');go.onclick=function(){var e=document.getElementById('settingsEntry');if(e)e.click();};box.append(go);
  $('Status').insertAdjacentElement('afterend',box);}
 $('Send').onclick=send;
 $('Cancel').onclick=function(){if(P.cancel)P.cancel();};
 $('Input').addEventListener('keydown',function(e){if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){e.preventDefault();send();}});
 $('Clear').onclick=function(){if(busy)return;history=[];status('');render();};
 root.AiChatPanel={history:function(){return history.slice();},send:send};
 render();
})(window);
