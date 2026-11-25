// AI 帮写公式：本机公式助手识别不了的需求，交给用户自己设置的 AI。
// 只发送表头、少量示例行和地址（不发整张表），经过隐私保护；返回的公式先预览，确认后只写入一个空白单元格。
(function(root){
 'use strict';
 if(!root.OfficeTools||root.OfficeTools.host()!=='et')return;
 var T=root.OfficeTools;
 function colName(n){var s='';while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
 var wpsHost={
  context:function(){var a=T.app(),ws=a.ActiveSheet,used=ws.UsedRange,cell=a.ActiveCell||a.Selection.Cells.Item(1,1),cols=Math.min(40,used.Columns.Count),c0=used.Column,headers=[],rows=[];
   for(var j=0;j<cols;j++){var v=used.Cells.Item(1,j+1).Value2;headers.push({col:colName(c0+j),name:v==null?'':String(v)});}
   for(var i=2;i<=Math.min(6,used.Rows.Count);i++){var row=[];for(var k=1;k<=cols;k++){var x=used.Cells.Item(i,k).Value2;row.push(x==null?'':x);}rows.push(row);}
   return {sheet:String(ws.Name),range:used.Address(false,false),target:cell.Address(false,false),headers:headers,rows:rows};},
  write:function(target,formula,sheet){var a=T.app(),ws=a.ActiveSheet;if(sheet&&String(ws.Name)!==sheet)throw Error('当前工作表已切换，请回到“'+sheet+'”再写入，或重新让 AI 生成。');var c=ws.Range(target);if(c.Cells.Count!==1)throw Error('结果位置只能是一个单元格。');
   if(c.HasFormula||(c.Value2!==undefined&&c.Value2!==null&&c.Value2!==''))throw Error('结果单元格 '+target+' 不是空的，请换一个空白单元格。');
   c.Formula=formula;var v=c.Value2;return typeof v==='string'&&/^#/.test(v)?'公式已写入 '+target+'，但结果显示 '+v+'，请检查引用或数据类型。':'公式已写入 '+target+'，结果：'+(v===undefined||v===null?'（空）':v)+'。';}
 };
 function H(){return root.AiFormulaHost||wpsHost;}
 var token='';
 async function api(body){
  if(!token)token=(await fetch('/api/session').then(function(r){return r.json();})).token;
  var send=function(b){return fetch('/api/ai/formula',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify(b)});};
  var r=await send(body);if(r.status===403){token=(await fetch('/api/session').then(function(x){return x.json();})).token;r=await send(body);}var d=await r.json().catch(function(){return {};});
  if(r.status===428&&d.needConfirm){if(!(await T.lib.ask(d.error)))throw Error('已取消，没有发送给 AI。');r=await send(Object.assign({},body,{privacyConfirmed:true}));d=await r.json().catch(function(){return {};});}
  if(!r.ok)throw Error(d.error||'AI 没有给出公式，请重试。');return d;
 }
 function mount(){
  var fa=document.getElementById('formulaAssistant');if(!fa||document.getElementById('aiFormula'))return !!fa;
  var box=document.createElement('div');box.id='aiFormula';box.className='ai-formula';
  box.innerHTML='<p class="input-hint">识别不了需求？先在表格里点一下要放结果的空白单元格，再交给 AI 写公式（只发送表头和前 5 行示例）。</p>'
   +'<button type="button" id="afGo">交给 AI 写公式</button><div id="afStatus" role="status" aria-live="polite"></div>'
   +'<div id="afResult" hidden><pre id="afFormula" translate="no"></pre><p class="input-hint" id="afExplain"></p><div class="button-pair"><button type="button" id="afApply" class="primary">写入该单元格</button><button type="button" id="afCopy">复制公式</button></div></div>';
  var anchor=fa.querySelector('#faRecognize');(anchor&&anchor.parentNode?anchor.parentNode:fa).insertBefore(box,anchor?anchor.nextSibling:null);
  var $=function(id){return box.querySelector('#af'+id);},got=null;
  function status(s,e){$('Status').textContent=s||'';$('Status').className='fa-status'+(e?' error':'');}
  $('Go').onclick=async function(){var intent=(document.getElementById('faIntent')||{}).value||'';intent=String(intent).trim();if(!intent){status('请先在“你想做什么？”里写下需求。',true);return;}
   $('Go').disabled=true;$('Result').hidden=true;got=null;status('AI 正在写公式…');
   try{var ctx=await H().context();var lang=root.OfficeI18n&&root.OfficeI18n.lang&&root.OfficeI18n.lang()==='ja'?'ja':'zh';var d=await api({intent:intent,context:ctx,lang:lang});
    got={formula:d.formula,target:ctx.target,sheet:ctx.sheet};$('Formula').textContent=d.formula;$('Explain').textContent=(d.explain||'')+' 结果位置：'+ctx.target;$('Result').hidden=false;
    status('请核对公式引用的列是否正确，再写入。'+(d.privacy&&d.privacy.masked?' 发送前隐藏了 '+d.privacy.masked+' 处个人信息。':''));}
   catch(e){status(e.message,true);}finally{$('Go').disabled=false;}};
  $('Apply').onclick=async function(){if(!got)return;try{if(!(await T.lib.ask('将把公式写入空白单元格 '+got.target+'：\n'+got.formula+'\n是否继续？')))return;status(await H().write(got.target,got.formula,got.sheet));got=null;$('Result').hidden=true;}catch(e){status(e.message,true);}};
  $('Copy').onclick=async function(){if(!got)return;try{await navigator.clipboard.writeText(got.formula);status('已复制公式。');}catch(e){status('复制失败，请手动选中公式复制。',true);}};
  return true;
 }
 if(!mount())document.addEventListener('DOMContentLoaded',mount);
 root.AiFormula={wpsHost:wpsHost,mount:mount};
})(window);
