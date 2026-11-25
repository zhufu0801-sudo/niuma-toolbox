// AI 翻译、润色与写作：选中文字 → 选择处理方式 → 预览结果 → 替换原文或插在后面。
// 写作（总结、续写、扩写、精简、整理成条目、改成书面语）只处理文字和演示里选中的文字。
// 文字、演示：处理选中的文字；表格：逐格处理选中的文字单元格（公式和数字不动），结果默认写到右侧一列。
// 内容通过本机服务发给用户自己设置的 AI，发送前经过隐私保护（个人信息自动隐藏）。
(function(root){
 'use strict';
 var T=root.OfficeTools;if(!T)return;
 var ACTIONS=[['auto','自动互译（中⇄日）'],['zh2ja','翻译成日文'],['ja2zh','翻译成中文'],['keigo','改成日文商务敬语'],['polishZh','中文润色'],['polishJa','日文润色'],['en','翻译成英文']];
 var WRITE=[['summary','总结要点'],['continue','续写'],['expand','扩写'],['shorten','精简改短'],['bullets','整理成条目'],['formal','改成正式书面语']];
 function isWrite(a){return WRITE.some(function(x){return x[0]===a;});}
 function t(x){return x;}
 // ---------- 读取与写回（WPS 默认实现；Office 版由 OfficeShim 提供 root.AiTextHost 覆盖） ----------
 var wpsHost={
  read:function(){var a=T.app(),h=T.host();
   if(h==='et'){var r=a.Selection;if(!r||!r.Cells)throw Error('请先在表格里选中要处理的单元格。');var n=r.Cells.Count;if(n>2000)throw Error('选区太大，请少选一些单元格（最多 200 个有文字的单元格）。');
    var items=[];for(var i=1;i<=r.Rows.Count;i++)for(var j=1;j<=r.Columns.Count;j++){var c=r.Cells.Item(i,j);if(c.HasFormula)continue;var v=c.Value2;if(typeof v==='string'&&v.trim())items.push({row:i,col:j,text:v});}
    if(!items.length)throw Error('选区里没有文字单元格（公式和数字不会处理）。');if(items.length>200)throw Error('一次最多处理 200 个文字单元格，请分批选。');
    return {kind:'cells',items:items,rows:r.Rows.Count,cols:r.Columns.Count,address:r.Address(false,false)};}
   if(h==='wpp'){var s=a.ActiveWindow&&a.ActiveWindow.Selection,tr=null;try{tr=s.TextRange;}catch(e){}if(!tr||!String(tr.Text||'').trim()){try{tr=s.ShapeRange.Item(1).TextFrame.TextRange;}catch(e){tr=null;}}
    if(!tr||!String(tr.Text||'').trim())throw Error('请先选中文本框，或在文本框里选中文字。');return {kind:'text',text:String(tr.Text).replace(/\r/g,'\n'),target:tr};}
   var sel=a.Selection,text=sel?String(sel.Text||''):'';if(!text.replace(/[\r\n\s]/g,''))throw Error('请先选中要处理的文字。');
   return {kind:'text',text:text.replace(/\r$/,'').replace(/\r/g,'\n'),target:sel.Range};},
  write:function(src,result,mode){var a=T.app(),h=T.host();
   if(src.kind==='cells'){var r=a.Selection;if(r.Address(false,false)!==src.address)throw Error('选区已改变，请重新读取。');
    var target=mode==='replace'?r:r.Offset(0,src.cols);
    if(mode!=='replace'&&!arguments[3]){var used=0;src.items.forEach(function(it){var v=target.Cells.Item(it.row,it.col).Value2;if(v!==undefined&&v!==null&&v!=='')used++;});if(used)return {confirm:'右侧区域有 '+used+' 个单元格已有内容，写入会覆盖。是否继续？'};}
    src.items.forEach(function(it,k){var cell=target.Cells.Item(it.row,it.col),v=result[k];cell.Value2=/^[=+\-@]/.test(v)?"'"+v:v;});
    return mode==='replace'?'已替换 '+src.items.length+' 个单元格（可按 Ctrl+Z 撤销）。':'已把结果写到右侧一列（'+src.items.length+' 个单元格），原文保留。';}
   var out=String(result);
   if(h==='wpp'){if(mode==='replace')src.target.Text=out;else src.target.InsertAfter('\r'+out);return mode==='replace'?'已替换选中的文字（可按 Ctrl+Z 撤销）。':'已把结果插在原文后面。';}
   if(mode==='replace')src.target.Text=out.replace(/\n/g,'\r');else src.target.InsertAfter('\r'+out.replace(/\n/g,'\r'));
   return mode==='replace'?'已替换选中的文字（可按 Ctrl+Z 撤销）。':'已把结果插在原文后面。';},
  // AI 对话的“插入文档”：文字插在光标处（有选中内容时插在后面）；表格从当前单元格向下写（单独一行公式按公式写入）；演示插在选中文字后，没有就加文本框。
  insert:function(text,confirmed){var a=T.app(),h=T.host(),out=String(text).replace(/\s+$/,'');if(!out)throw Error('没有可插入的内容。');
   if(h==='et'){var cell=a.ActiveCell;if(!cell)throw Error('请先点一个单元格。');var lines=out.split('\n');if(lines.length>500)throw Error('内容超过 500 行，请复制后手动粘贴。');var formula=lines.length===1&&/^=/.test(lines[0].trim()),at=cell.Address(false,false),target=cell.Resize(lines.length,1);
    if(!confirmed){var used=0;for(var i=1;i<=lines.length;i++){var v=target.Cells.Item(i,1).Value2;if(v!==undefined&&v!==null&&v!=='')used++;}return {confirm:(formula?'将把公式写入 '+at:'将从 '+at+' 开始向下写入 '+lines.length+' 行')+(used?'，其中 '+used+' 个单元格已有内容，会被覆盖':'')+'。\nWPS 表格的撤销不能还原这一步。是否继续？'};}
    if(formula){cell.Formula=lines[0].trim();return '已把公式写入 '+at+'。';}
    lines.forEach(function(l,k){target.Cells.Item(k+1,1).Value2=/^[=+\-@]/.test(l)?"'"+l:l;});return '已从 '+at+' 开始写入 '+lines.length+' 行。';}
   if(h==='wpp'){var s=a.ActiveWindow&&a.ActiveWindow.Selection,tr=null;try{tr=s.TextRange;if(tr&&typeof tr.Text!=='string')tr=null;}catch(e){tr=null;}
    if(tr){tr.InsertAfter('\r'+out.replace(/\n/g,'\r'));return '已插在选中的文字后面。';}
    var slide=a.ActiveWindow.View.Slide;if(!slide)throw Error('请先点一下要插入的那一页。');var box=slide.Shapes.AddTextbox(1,60,60,600,300);box.TextFrame.TextRange.Text=out.replace(/\n/g,'\r');return '已在当前页加了一个文本框（可以拖动位置）。';}
   var r=a.Selection&&a.Selection.Range;if(!r)throw Error('请先把光标放到要插入的位置。');var txt=out.replace(/\n/g,'\r');if(Number(r.End)>Number(r.Start))r.InsertAfter('\r'+txt);else r.InsertAfter(txt);return '已插入到文档（可按 Ctrl+Z 撤销）。';}
 };
 function H(){return root.AiTextHost||wpsHost;}
 var token='';
 var current=null,cancelled=false,OLD='本机服务还是旧版本：请双击桌面“Office Toolbox”重启后再试。';
 function cancel(){if(current){cancelled=true;current.abort();}}
 // 等待时每秒更新：已等待几秒；本机模型第一次加载、会先“思考”的模型要更久。
 function waiting(set,base){var t0=Date.now(),local=/Ollama|LM Studio|本机/.test(($('Note')&&$('Note').textContent)||'');set(base+(local?'本机模型第一次加载可能要 1～3 分钟。':'通常需要 5～60 秒。'));var iv=setInterval(function(){set(base+'已等待 '+Math.round((Date.now()-t0)/1000)+' 秒'+(local?'（本机模型第一次加载可能要 1～3 分钟）':'')+'，可以点“取消”。');},1000);return function(){clearInterval(iv);};}
 async function api(url,body){
  if(!token){var s=await fetch('/api/session').then(function(r){return r.json();});token=s.token;}
  // 每次请求一个 AbortController：点“取消”或超过 310 秒都会中止。
  var ctl=typeof AbortController!=='undefined'?new AbortController():null,timer=ctl?setTimeout(function(){ctl.abort();},310000):0;current=ctl;
  var send=function(b){return fetch(url,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify(b),signal:ctl?ctl.signal:undefined});};
  var r,d;try{r=await send(body);if(r.status===403){token=(await fetch('/api/session').then(function(x){return x.json();})).token;r=await send(body);}d=await r.json().catch(function(){return {};});}
  catch(e){if(e&&e.name==='AbortError')throw Error(cancelled?'已取消。':'AI 太久没有回应（超过 5 分钟），已停止。可以换一个更快的模型再试。');if(r&&r.status===404)throw Error(OLD);throw Error('连不上本机服务：请双击桌面“Office Toolbox”后再试。');}
  finally{clearTimeout(timer);if(current===ctl)current=null;cancelled=false;}
  if(r.status===404)throw Error(OLD);
  if(r.status===428&&d.needConfirm){if(!(await T.lib.ask(d.error)))throw Error('已取消，没有发送给 AI。');r=await send(Object.assign({},body,{privacyConfirmed:true}));d=await r.json().catch(function(){return {};});}
  if(!r.ok)throw Error(d.error||'AI 没有完成处理，请重试。');return d;
 }
 var section=document.createElement('section');section.className='tool-card ai-text-panel';section.dataset.panel='quick';section.id='aiTextPanel';
 section.innerHTML='<details><summary><strong>AI 翻译、润色与写作</strong><span>中⇄日 · 敬语 · 总结 · 续写 · 扩写 · 先预览再写入</span></summary>'
  +'<p class="input-hint" id="atNote">选中文字后点“读取选区”。</p>'
  +'<label for="atAction">处理方式</label><select id="atAction"></select>'
  +'<div id="atExtraBox" hidden><label for="atExtra">补充要求（可不填）</label><input id="atExtra" maxlength="300" autocomplete="off" placeholder="例如：控制在 200 字以内、写给客户"></div>'
  +'<button type="button" id="atRead">读取选区</button><p class="input-hint" id="atSource"></p>'
  +'<button type="button" id="atGo" class="primary" disabled>让 AI 处理</button><button type="button" id="atCancel" hidden>取消</button>'
  +'<div id="atStatus" role="status" aria-live="polite"></div>'
  +'<div id="atResult" hidden><label for="atOut">结果（可以先改再写入）</label><textarea id="atOut" rows="6"></textarea>'
  +'<div class="button-pair"><button type="button" id="atInsert">插在原文后</button><button type="button" id="atReplace">替换原文</button></div></div></details>';
 var $=function(id){return section.querySelector('#at'+id);},src=null,result=null;
 function place(){var office=document.getElementById('toolSection');if(office&&office.parentNode&&!document.getElementById('replaceSection')){office.parentNode.insertBefore(section,office);return true;}var anchor=document.getElementById('replaceSection');if(anchor&&anchor.parentNode){anchor.parentNode.insertBefore(section,anchor);return true;}var q=document.querySelector('.quick-section');if(q&&q.parentNode){q.parentNode.insertBefore(section,q.nextSibling);return true;}return false;}
 if(!place()){document.addEventListener('DOMContentLoaded',place);}
 [['翻译与润色',ACTIONS],['写作',WRITE]].forEach(function(g){var og=document.createElement('optgroup');og.label=g[0];g[1].forEach(function(x){var o=document.createElement('option');o.value=x[0];o.textContent=x[1];og.appendChild(o);});$('Action').appendChild(og);});
 try{var lastAction=localStorage.getItem('officeToolbox.aiText.action');if(lastAction)$('Action').value=lastAction;}catch(e){}
 function status(s,err){$('Status').textContent=s||'';$('Status').className='fa-status'+(err?' error':'');}
 function cells(){return src&&src.kind==='cells';}
 function sync(){var a=$('Action').value;$('ExtraBox').hidden=!isWrite(a);$('Go').disabled=!src;$('Result').hidden=!result;$('Insert').textContent=cells()?'写到右侧一列':a==='continue'?'接在原文后':'插在原文后';$('Replace').hidden=a==='continue';$('Out').readOnly=cells();}
 // 首次说明：内容会发送给用户选择的服务商
 fetch('/api/ai/config').then(function(r){return r.json();}).then(function(d){var c=d.config||{},p=(d.providers||[]).find(function(x){return x.id===c.provider;});
   if(!c.ready){$('Note').textContent='还没有接入 AI：请到“设置 → AI 服务”填写服务商和 API Key。';return;}
   var mode=(d.privacy&&d.privacy.mode)||'mask';var ja=root.OfficeI18n&&root.OfficeI18n.lang&&root.OfficeI18n.lang()==='ja';$('Note').textContent='将使用 '+(p?(ja&&p.nameJa?p.nameJa:p.name):c.provider)+'。'+(p&&p.region==='local'?'内容只在这台电脑上处理。':'选中的文字会发送给该服务商；'+({mask:'身份证号、手机号等会先自动隐藏。',remind:'含个人信息时会先问你。',off:'隐私保护已关闭。'}[mode]));}).catch(function(){});
 $('Action').onchange=function(){result=null;try{localStorage.setItem('officeToolbox.aiText.action',$('Action').value);}catch(e){}sync();};
 $('Read').onclick=async function(){result=null;try{src=await H().read();$('Source').textContent=cells()?'已读取 '+src.address+' 中的 '+src.items.length+' 个文字单元格':'已读取 '+Array.from(src.text).length+' 个字：'+src.text.slice(0,60)+(src.text.length>60?'…':'');status('');}catch(e){src=null;$('Source').textContent='';status(e.message,true);}sync();};
 $('Cancel').onclick=cancel;
 $('Go').onclick=async function(){if(!src)return;$('Go').disabled=true;$('Cancel').hidden=false;var stop=waiting(function(t){status(t);},'AI 正在处理，');
  try{var lang=(root.OfficeI18n&&root.OfficeI18n.lang&&root.OfficeI18n.lang()==='ja')?'ja':'zh',act=$('Action').value,body={action:act,lang:lang},write=isWrite(act);
   if(write&&cells())throw Error('总结、续写、扩写等写作功能请在文字或演示里选中一段文字；表格里可以用翻译和润色。');
   if(cells())body.items=src.items.map(function(x){return x.text;});else body.text=src.text;if(write)body.extra=$('Extra').value.trim();
   var d=await api(write?'/api/ai/write':'/api/ai/translate',body);result=cells()?d.items:d.text;
   $('Out').value=cells()?src.items.map(function(x,k){return x.text+'  →  '+result[k];}).join('\n'):result;
   status('已完成（'+d.model+'）。'+(d.privacy&&d.privacy.masked?'发送前隐藏了 '+d.privacy.masked+' 处个人信息，结果里已还原。':'')+'请检查后再写入。');}
  catch(e){result=null;status(e.message,true);}finally{stop();$('Cancel').hidden=true;}sync();};
 async function write(mode,confirmed){try{var value=cells()?result:$('Out').value;var r=await H().write(src,value,mode,confirmed);
   if(r&&r.confirm){if(!(await T.lib.ask(r.confirm)))return;r=await H().write(src,value,mode,true);}
   status(r);result=null;src=null;$('Source').textContent='';sync();}catch(e){status(e.message,true);}}
 $('Insert').onclick=function(){write('insert');};
 $('Replace').onclick=async function(){if(cells()&&!(await T.lib.ask('将用结果覆盖选中的 '+src.items.length+' 个单元格（可按 Ctrl+Z 撤销）。是否继续？')))return;write('replace');};
 function select(action){var ok=ACTIONS.concat(WRITE).some(function(x){return x[0]===action;});if(!ok)return false;$('Action').value=action;result=null;try{localStorage.setItem('officeToolbox.aiText.action',action);}catch(e){}sync();var dt=section.querySelector('details');if(dt)dt.open=true;return true;}
 root.AiTextPanel={wpsHost:wpsHost,host:H,api:api,select:select,cancel:cancel,waiting:waiting};
 sync();
})(window);
