// Explicit copy/paste only. No clipboard watcher and no remote AI calls.
(function(root){'use strict';
function wps(){var t=root.OfficeTools;return t&&t.app&&t.app()?t:null;}
function officeCall(method,value){return new Promise(function(resolve,reject){var o=root.Office;if(!o||!o.context||!o.context.document)return reject(Error('Office API unavailable'));var cb=function(r){if(r.status==='succeeded')resolve(r.value);else reject(Error(r.error&&r.error.message||'Office API unavailable'));};if(method==='read')o.context.document.getSelectedDataAsync(o.CoercionType.Text,cb);else o.context.document.setSelectedDataAsync(value,{coercionType:o.CoercionType.Text},cb);});}
async function selection(){var t=wps();if(t){var a=t.app(),h=t.host();if(h==='et'){var r=a.Selection,n=Number(r.Rows.Count),m=Number(r.Columns.Count),rows=[];if(n*m>20000)throw Error('20,000 cells / セル / 单元格');for(var i=1;i<=n;i++){var row=[];for(var j=1;j<=m;j++)row.push(String(r.Cells.Item(i,j).Text||''));rows.push(row.join('\t'));}return rows.join('\n');}if(h==='wpp')return String(a.ActiveWindow.Selection.TextRange.Text||'');return String(a.Selection.Text||'').replace(/\r/g,'\n').replace(/\u0007/g,'\t');}
 if(root.Excel&&root.Office&&/Excel/i.test(root.Office.context.host))return root.Excel.run(async function(c){var r=c.workbook.getSelectedRange();r.load('rowCount,columnCount');await c.sync();if(r.rowCount*r.columnCount>20000)throw Error('20,000 cells / セル / 单元格');r.load('text');await c.sync();return r.text.map(function(row){return row.join('\t');}).join('\n');});
 return String(await officeCall('read')||'');}
async function insert(text){var t=wps();if(t){var a=t.app(),h=t.host();if(h==='et'){a.ActiveCell.NumberFormat='@';a.ActiveCell.Value2=text;return;}if(h==='wpp'){a.ActiveWindow.Selection.TextRange.Text=text;return;}a.Selection.TypeText(text.replace(/\n/g,'\r'));return;}return officeCall('write',text);}
function bounded(p,ms){return new Promise(function(resolve,reject){var timer=setTimeout(function(){reject(Error('Clipboard timeout'));},ms);Promise.resolve(p).then(function(v){clearTimeout(timer);resolve(v);},function(e){clearTimeout(timer);reject(e);});});}
async function copyText(text,env){env=env||root;try{if(env.navigator.clipboard&&env.navigator.clipboard.writeText){await bounded(env.navigator.clipboard.writeText(text),2000);return true;}}catch(ignore){}
 var ta=env.document.createElement('textarea');ta.value=text;ta.style.position='fixed';ta.style.opacity='0';env.document.body.appendChild(ta);var ok=false;try{ta.focus();ta.select();ok=!!env.document.execCommand('copy');}catch(ignore){}finally{ta.remove();}return ok;}
async function readText(env,timeout){try{env=env||root;if(env.navigator.clipboard&&env.navigator.clipboard.readText)return await bounded(env.navigator.clipboard.readText(),timeout||1500);}catch(ignore){}return null;}
function mount(container,o){var E=root.MaskEngine,M=o.M,el=o.el,post=o.post,say=o.say,run=o.task,restored='';
 var panel=el('fieldset','mk-group');panel.id='maskClipboard';panel.append(el('legend','',M('给网页 AI 用','Web AI とのコピー・貼り付け')));
 var bar=el('div','button-pair'),copy=el('button','',M('脱敏复制','マスキングしてコピー')),paste=el('button','',M('还原粘贴','元に戻して貼り付け'));copy.type=paste.type='button';copy.id='mkClipCopy';paste.id='mkClipPaste';bar.append(copy,paste);
 var note=el('p','mk-note',M('使用当前勾选的种类和词表，固定换成代号。只保留最近 20 次复制的对应表；同一代号以最新记录为准。','今選んでいる種類と単語リストを使い、記号に置き換えます。対応表は最近のコピー 20 回分を保存し、同じ記号は最新の記録を使います。'));
 var input=el('textarea');input.id='mkClipInput';input.rows=4;input.placeholder=M('把 AI 的回答粘贴到这里，再点“还原并插入”','AI の回答をここに貼り付けて「復元して挿入」を押してください');input.setAttribute('aria-label',M('AI 回答','AI の回答'));
 var manual=el('button','',M('还原并插入','復元して挿入')),copyBack=el('button','',M('复制还原后的文字','復元した文字をコピー'));manual.type=copyBack.type='button';
 var output=el('textarea');output.id='mkClipOutput';output.rows=4;output.readOnly=true;output.hidden=true;output.setAttribute('aria-label',M('复制或还原结果','コピー・復元結果'));
 function show(text){output.value=text;output.hidden=false;}
 function report(r){say(M('已还原 '+r.done+' 处。','復元しました：'+r.done+' か所。')+(r.unknown.length?M(' 未知代号保留：',' 対応表にない記号はそのまま残しました：')+r.unknown.join('、'):''));}
 async function restore(text){if(!String(text||'').trim())throw Error(M('请先粘贴 AI 的回答。','先に AI の回答を貼り付けてください。'));var d=await post({action:'clip-pairs'});return E.restoreCodes(text,d.pairs);}
 async function pasteText(text){var r=await restore(text);restored=r.text;show(restored);await insert(restored);report(r);}
 copy.onclick=function(){run(async function(){var text=await selection();if(!text.trim())throw Error(M('请先在文档里选中文字或单元格。','先に文書の文字かセルを選択してください。'));if(text.length>200000)throw Error(M('一次最多复制 20 万字符。','一度にコピーできるのは 20 万文字までです。'));var opts=o.opts(),sess=E.session({mode:'code',lang:o.lang(),words:opts.words}),r=E.maskText(text,sess,opts),counts={};r.seq.forEach(function(p){counts[p[2]]=(counts[p[2]]||0)+1;});
  await post({action:'save',record:{kind:'clip',mode:'code',host:wps()?wps().host():'office-word',name:'复制给 AI',pairs:E.table([r.seq]),counts:counts,total:r.seq.length}});await o.refresh();var ok=await copyText(r.text);show(r.text);say(ok?M('已复制脱敏文字。','マスキングした文字をコピーしました。'):M('不能自动复制。请选中下面的结果，按 Ctrl+C。','自動でコピーできません。下の結果を選択して Ctrl+C を押してください。'));});};
 paste.onclick=function(){run(async function(){var text=await readText();if(text===null){input.hidden=false;input.focus();say(M('读不到剪贴板。请把回答粘到下面，再点“还原并插入”。','クリップボードを読み取れません。下の欄に回答を貼り付けて「復元して挿入」を押してください。'));return;}input.value=text;await pasteText(text);});};
 manual.onclick=function(){run(function(){return pasteText(input.value);});};
 copyBack.onclick=function(){run(async function(){var r=await restore(input.value);restored=r.text;show(restored);var ok=await copyText(restored);report(r);if(!ok)say(M('请选中下面的还原结果，按 Ctrl+C。','下の復元結果を選択して Ctrl+C を押してください。'));});};
 panel.append(bar,note,input,manual,copyBack,output);container.append(panel);
}
root.MaskClipboard={mount:mount,selection:selection,insert:insert,copyText:copyText,readText:readText};
})(window);
