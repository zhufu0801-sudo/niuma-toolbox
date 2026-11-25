(function(root){'use strict';
// 粘贴栏（侧栏“文件与内容”页最上面）：复制过、收集过的文字和表格排成一列，点“插入”放进当前文档。
// 手动：选中内容点“收集选中内容”，或点“从剪贴板添加”。自动：勾选“自动记录复制的文字”（默认关闭）。
var T=root.OfficeTools,S=root.WorkbenchSettings,box=null,list=null,status=null,toggle=null,skipNote=null,signature='',busy=false,timer=null;
function t(s){return root.OfficeI18n?root.OfficeI18n.t(s):s;}
function el(tag,cls,text){var e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;}
function say(message,bad){if(!status)return;status.textContent=message;status.className='pb-status'+(bad?' bad':'');}
function plain(block){return block.type==='table'?block.rows.map(function(r){return r.join('\t');}).join('\n'):block.text;}
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function when(at,now){var d=new Date(at),m=Math.round(((now||Date.now())-at)/60000),J=ja(),today=new Date(now||Date.now());if(m<1)return J?'たった今':'刚才';if(m<60)return m+(J?' 分前':' 分钟前');var hm=('0'+d.getHours()).slice(-2)+':'+('0'+d.getMinutes()).slice(-2);if(d.toDateString()===today.toDateString())return (J?'今日 ':'今天 ')+hm;return (d.getMonth()+1)+(J?'月':' 月 ')+d.getDate()+(J?'日':' 日');}
var FROM={wps:'来自文字',et:'来自表格',wpp:'来自演示',clipboard:'复制'};
function copyText(text){if(root.navigator&&navigator.clipboard&&navigator.clipboard.writeText)return navigator.clipboard.writeText(text).catch(fallback);return Promise.resolve().then(fallback);
  function fallback(){var a=document.createElement('textarea');a.value=text;a.setAttribute('readonly','');a.style.cssText='position:fixed;left:-9999px;top:0';document.body.append(a);a.select();var ok=false;try{ok=document.execCommand('copy');}catch(ignore){}a.remove();if(!ok)throw Error('这个窗口不能写入剪贴板，请改用“插入”。');}}
function item(x){var row=el('li','pb-item'+(x.pinned?' pinned':'')),preview=el('div','pb-text'),meta=el('div','pb-meta'),actions=el('div','pb-actions');
  if(x.block.type==='table'){preview.textContent=t('表格')+' '+x.block.rows.length+' × '+x.block.rows[0].length+'：'+x.block.rows.slice(0,2).map(function(r){return r.join(' | ');}).join(' / ');}else preview.textContent=x.block.text.slice(0,200);
  preview.setAttribute('translate','no');preview.title=plain(x.block).slice(0,600);
  meta.setAttribute('translate','no');meta.textContent=(x.pinned?'📌 ':'')+t(FROM[x.sourceHost]||'')+' · '+when(x.at);
  var ins=el('button','pb-insert','插入'),copy=el('button','','复制'),pin=el('button','',x.pinned?'取消钉住':'钉住'),del=el('button','pb-del','删除');
  [ins,copy,pin,del].forEach(function(b){b.type='button';});
  ins.title='放进当前文档的光标处，格式跟随周围文字';
  ins.onclick=function(){if(busy)return;busy=true;ins.disabled=true;Promise.resolve().then(function(){return T.runFromUi('pasteBarInsert',{block:x.block});}).then(function(m){say(m);},function(e){say('未完成：'+e.message,true);}).then(function(){busy=false;ins.disabled=false;});};
  copy.onclick=function(){copyText(plain(x.block)).then(function(){say('已复制，可以按 Ctrl+V 粘贴到任何地方。');},function(e){say(e.message,true);});};
  pin.onclick=function(){T.pasteBar.pin(x.id,!x.pinned).then(function(d){draw(d);},function(e){say(e.message,true);});};
  del.onclick=function(){T.pasteBar.remove(x.id).then(function(d){draw(d);},function(e){say(e.message,true);});};
  actions.append(ins,copy,pin,del);row.append(preview,meta,actions);return row;}
function draw(d){if(!list)return;var sig=JSON.stringify([d.items,d.watching,d.skipped,ja(),Math.floor(Date.now()/60000)]);if(sig===signature)return;signature=sig;list.textContent='';
  (d.items||[]).forEach(function(x){list.append(item(x));});box.querySelector('.pb-empty').hidden=(d.items||[]).length>0;box.querySelector('.pb-clear').hidden=!(d.items||[]).some(function(x){return !x.pinned;});
  skipNote.hidden=!(d.watching&&d.skipped>0);skipNote.textContent=t('为了安全，已跳过 ')+d.skipped+t(' 条像密码、证件号或卡号的内容。');}
function refresh(){if(!box||box.hidden||!T.pasteBar||document.hidden)return;T.pasteBar.list().then(draw,function(){});}
function watching(){try{var s=S&&S.snapshot();return !!(s&&s.settings&&s.settings.pasteBar&&s.settings.pasteBar.watch);}catch(e){return false;}}
function build(){var files=document.querySelector('section[data-panel="files"]');if(!files||!T||!T.pasteBar)return;
  box=el('section','tool-card pastebar');box.id='pasteBarSection';box.dataset.panel='files';box.hidden=files.hidden;
  var head=el('div','card-heading'),icon=el('span','section-icon teal'),img=el('img');img.src='icons/insert.svg';img.alt='';icon.append(img);var h=el('div');h.append(el('h2','','粘贴栏'),el('p','','复制过、收集过的内容排在这里，点“插入”放进当前文档'));head.append(icon,h);
  var bar=el('div','button-pair'),collect=el('button','','收集选中内容'),fromClip=el('button','','从剪贴板添加');collect.type=fromClip.type='button';
  collect.title='把当前选中的文字或表格放进粘贴栏，原文不变';fromClip.title='把刚才复制（Ctrl+C）的文字放进粘贴栏';
  collect.onclick=function(){T.runFromUi('pasteBarCollect').then(function(m){say(m);signature='';refresh();},function(e){say('未完成：'+e.message,true);});};
  fromClip.onclick=function(){T.pasteBar.fromClipboard().then(function(d){say('已从剪贴板添加。');draw(d);},function(e){say(e.message,true);});};
  bar.append(collect,fromClip);
  var label=el('label','pb-watch');toggle=el('input');toggle.type='checkbox';toggle.checked=watching();label.append(toggle,el('span','','自动记录复制的文字'));
  var note=el('p','pb-note','只存在这台电脑上，最多 30 条（钉住的另算）；会自动跳过像密码、身份证号、银行卡号的内容。');
  toggle.onchange=function(){var on=toggle.checked;if(!S){toggle.checked=!on;return;}S.update({pasteBar:{watch:on}}).then(function(){say(on?'已开始记录：之后复制的文字会出现在这里。':'已停止自动记录。');signature='';refresh();},function(e){toggle.checked=!on;say(e.message,true);});};
  if(S&&S.subscribe)S.subscribe(function(){if(toggle)toggle.checked=watching();});
  skipNote=el('p','pb-note');skipNote.setAttribute('translate','no');skipNote.hidden=true;status=el('p','pb-status');status.setAttribute('role','status');
  list=el('ul','pb-list');var empty=el('p','pb-empty','还没有内容。选中文字后点“收集选中内容”，或者复制后点“从剪贴板添加”。');
  var clear=el('button','link-button pb-clear','清空（保留钉住的）');clear.type='button';clear.hidden=true;clear.onclick=function(){if(!root.confirm(t('清空粘贴栏里没钉住的内容？')))return;T.pasteBar.clear().then(function(d){say('已清空。');draw(d);},function(e){say(e.message,true);});};
  box.append(head,bar,label,note,skipNote,status,list,empty,clear);files.parentNode.insertBefore(box,files);
  // 切换到“文件与内容”页时，粘贴栏跟着显示 / 隐藏（侧栏按 data-panel 切换）。
  if(typeof MutationObserver==='function')new MutationObserver(function(){if(!box.hidden){signature='';refresh();}}).observe(box,{attributes:true,attributeFilter:['hidden']});
  timer=setInterval(refresh,2000);refresh();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',build);else build();
root.PasteBar={refresh:refresh,when:when,plain:plain};
})(window);
