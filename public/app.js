(function () {
  'use strict';
  var labels={wps:'文字',et:'表格',wpp:'PPT'},host=new URLSearchParams(location.search).get('host')||'wps';
  if(!labels[host])host='wps';
  var context={},history=[],token='',busy=false,lastContext='';
  var $=function(id){return document.getElementById(id);};
  function setHost(value){host=value;document.querySelectorAll('[data-host]').forEach(function(b){b.classList.toggle('active',b.dataset.host===host);b.setAttribute('aria-pressed',String(b.dataset.host===host));});renderSuggestions();refreshContext();$('question').placeholder={wps:'例如：怎么给这段文字设置标题并生成目录？',et:'例如：怎么按部门汇总这些数据？',wpp:'例如：这页 PPT 的文字太多，该怎么排版？'}[host];}
  function refreshContext(){
    context=window.CoachBridge.read(host);var signature=JSON.stringify(context);if(signature===lastContext)return;lastContext=signature;
    $('previewNote').hidden=context.connected;
    $('contextTitle').textContent=context.document || '在 WPS 中打开文档';
    $('contextDetail').textContent=(context.location ? context.location+' · ' : '')+(context.detail||'这里也可以直接学习操作，文档选区只在 WPS 侧边栏内读取。');
    $('contextPreview').textContent=context.selection||'暂未读取到选中内容。';
    $('suggestionTitle').textContent=context.selection ? '你接下来可能想做' : '从一个小操作开始';
  }
  function renderSuggestions(){
    $('suggestions').replaceChildren();
    window.CoachGuides[host].forEach(function(guide){
      var b=document.createElement('button');b.type='button';b.className='suggestion';
      var icon=document.createElement('span');icon.className='suggestion-icon';icon.textContent=guide.icon;
      var copy=document.createElement('span'),title=document.createElement('span'),desc=document.createElement('span'),arrow=document.createElement('span');
      title.className='suggestion-title';title.textContent=guide.title;desc.className='suggestion-desc';desc.textContent=guide.desc;arrow.className='suggestion-arrow';arrow.textContent='↗';copy.append(title,desc);b.append(icon,copy,arrow);
      b.addEventListener('click',function(){addMessage('local',guide.answer,'操作小课堂 · '+labels[host]);$('question').value=guide.question;$('question').focus();});$('suggestions').appendChild(b);
    });
  }
  function addMessage(role,content,label){
    $('emptyState').hidden=true;
    var item=document.createElement('article');item.className='message '+(role==='local'?'assistant local':role);
    var heading=document.createElement('div');heading.className='message-label';var name=document.createElement('span');name.textContent=label||(role==='user'?'你':'DeepSeek · 本机教练');heading.appendChild(name);
    if(role!=='user'){var copy=document.createElement('button');copy.type='button';copy.className='message-copy';copy.textContent='复制';copy.addEventListener('click',async function(){try{await navigator.clipboard.writeText(content);copy.textContent='已复制';}catch{$('question').value=content;copy.textContent='已放入输入框';}});heading.appendChild(copy);}
    var body=document.createElement('div');body.className='message-body';
    String(content).split(/(\*\*[^*\n]+\*\*|`[^`\n]+`)/g).forEach(function(part){
      if(part.slice(0,2)==='**'&&part.slice(-2)==='**'){var strong=document.createElement('strong');strong.textContent=part.slice(2,-2);body.appendChild(strong);}
      else if(part[0]==='`'&&part.slice(-1)==='`'){var code=document.createElement('code');code.textContent=part.slice(1,-1);body.appendChild(code);}
      else body.appendChild(document.createTextNode(part));
    });
    item.append(heading,body);$('messages').appendChild(item);item.scrollIntoView({block:'nearest',behavior:'smooth'});return item;
  }
  async function session(){var response=await fetch('/api/session');if(!response.ok)throw new Error('本地服务未连接。');token=(await response.json()).token;}
  async function health(){
    try{var response=await fetch('/api/health');var data=await response.json();var ready=data.engine&&data.installed;$('connection').className='connection '+(ready?'ready':'error');$('modelStatus').textContent=ready?'DeepSeek 已就绪 · 数据留在本机':!data.engine?'请先启动 Ollama':'DeepSeek 模型尚在下载';}catch{$('connection').className='connection error';$('modelStatus').textContent='本地服务未连接，请重新启动教练';}
  }
  $('askForm').addEventListener('submit',async function(e){
    e.preventDefault();if(busy)return;var question=$('question').value.trim();if(!question)return;
    busy=true;$('sendButton').disabled=true;$('sendButton').textContent='正在整理步骤…';refreshContext();
    var requestContext=$('includeContext').checked?{document:context.document,selection:context.selection,location:context.location,detail:context.detail}:{};
    addMessage('user',question);$('question').value='';
    var waiting=addMessage('assistant','正在结合你的问题整理操作步骤。首次使用时，需要先加载本机模型。');
    try{
      if(!token)await session();
      var ask=function(){return fetch('/api/ask',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify({host:host,question:question,context:requestContext,history:history})});};var response=await ask();if(response.status===403){await session();response=await ask();}var result=await response.json();waiting.remove();
      if(!response.ok)throw new Error(result.error||'本机模型未能完成这次回答。');
      addMessage('assistant',result.answer);history.push({role:'user',content:question},{role:'assistant',content:result.answer});history=history.slice(-6);
    }catch(error){waiting.remove();addMessage('error',error.message,'暂时未能回答');}
    finally{busy=false;$('sendButton').disabled=false;$('sendButton').textContent='开始指导 ↗';health();}
  });
  $('question').addEventListener('keydown',function(e){if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){e.preventDefault();$('askForm').requestSubmit();}});
  $('clearChat').addEventListener('click',function(){if(busy)return;history=[];$('messages').replaceChildren();$('emptyState').hidden=false;});
  $('refreshContext').addEventListener('click',function(){lastContext='';refreshContext();});$('recheck').addEventListener('click',health);
  document.querySelectorAll('[data-host]').forEach(function(b){b.addEventListener('click',function(){if(busy)return;if(context.connected&&b.dataset.host!==host){addMessage('local','请在对应的 WPS 文字、表格或演示窗口里打开教练。当前仍使用这份文档的操作指导。');return;}history=[];setHost(b.dataset.host);});});
  setHost(host);session().catch(function(){});health();setInterval(refreshContext,1800);
}());
