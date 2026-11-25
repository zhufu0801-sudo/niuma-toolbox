// 办公工具箱 · 宏代码检查侧栏（2026-11-25）。在侧栏“文件”页（Office 版在工具列表下面）。
// 代码可以粘贴进来，也可以读当前文件里的宏（要先在宏安全设置里勾选“信任对 VBA 工程的访问”）。
// 只看写法、不运行代码；能确定怎么改的，生成修改后的代码给客户复制，不直接写回文件。
(function(root){'use strict';
var V=root.VbaLint,box=null,modules=[],results=[],busy=false;
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
function el(tag,cls,text){var e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;}
function $(id){return document.getElementById(id);}
function say(m,bad){var s=$('vbStatus');if(!s)return;s.textContent=m||'';s.className='mk-status'+(bad?' bad':'');}
function wps(){var T=root.OfficeTools;return T&&T.app&&T.app()?T:null;}
// ---- 读取当前文件里的宏 ----
function activeDoc(){var T=wps();if(!T)return null;var a=T.app(),h=T.host();return h==='et'?a.ActiveWorkbook:h==='wpp'?a.ActivePresentation:a.ActiveDocument;}
var TRUST=function(){return M('读不到宏代码。请先打开“信任对 VBA 工程的访问”：\nWPS：工具 → 宏 → 安全性 → 可靠发行商 → 勾选“信任对于 Visual Basic 项目的访问”。\nOffice：文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 → 勾选“信任对 VBA 工程对象模型的访问”。\n改完后重新打开文件。也可以直接把代码复制粘贴到下面的框里。','マクロのコードを読み取れません。先に「VBA プロジェクトへのアクセスを信頼する」を有効にしてください：\nWPS：ツール → マクロ → セキュリティ → 信頼できる発行元 → 「Visual Basic プロジェクトへのアクセスを信頼する」にチェック。\nOffice：ファイル → オプション → トラスト センター → トラスト センターの設定 → マクロの設定 → 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック。\n変更後、ファイルを開き直してください。コードを下の欄に貼り付けてもチェックできます。');};
function readMacros(){if(!wps())throw Error(M('Office 版读不到宏代码（Office 的加载项接口不支持）。请在 VBA 编辑器（Alt+F11）里复制代码，粘贴到下面的框里。','Office 版ではマクロのコードを読み取れません（Office アドインの仕組みで未対応）。VBA エディター（Alt+F11）でコードをコピーし、下の欄に貼り付けてください。'));
 var d=activeDoc();if(!d)throw Error(M('请先打开一个文件。','先にファイルを開いてください。'));var comps;try{comps=d.VBProject.VBComponents;var c=Number(comps.Count);if(!(c>=0))throw Error('x');}catch(e){throw Error(TRUST());}
 var list=[];for(var i=1;i<=Number(comps.Count);i++){try{var m=comps.Item(i),cm=m.CodeModule,n=Number(cm.CountOfLines);if(n>0)list.push({name:String(m.Name),code:String(cm.Lines(1,n))});}catch(ignore){}}
 if(!list.length)throw Error(M('这个文件里没有宏代码。','このファイルにはマクロのコードがありません。'));return list;}
// ---- 检查 ----
function run(){var code=$('vbCode').value;var sel=$('vbModule');var list=modules.length&&sel&&sel.value!=='__paste'?modules.filter(function(m){return sel.value==='__all'||m.name===sel.value;}):[{name:'',code:code}];
 if(!list.some(function(m){return m.code.trim();}))throw Error(M('请先粘贴 VBA 代码，或者点“读取当前文件的宏”。','VBA のコードを貼り付けるか、「今のファイルのマクロを読み取る」を押してください。'));
 results=list.map(function(m){var r=V.lint(m.code);return {name:m.name,code:m.code,r:r};});render();}
function sevName(s){return M(V.LABEL[s][0],V.LABEL[s][1]);}
function render(){var wrap=$('vbResult');wrap.textContent='';var total={error:0,risk:0,warn:0,tip:0},any=false;results.forEach(function(x){Object.keys(total).forEach(function(k){total[k]+=x.r.count[k];});});
 var chips=el('div','vb-chips');['error','risk','warn','tip'].forEach(function(k){var c=el('span','vb-chip vb-'+k,sevName(k)+' '+total[k]);chips.append(c);});wrap.append(chips);
 results.forEach(function(x){if(!x.r.findings.length)return;any=true;if(x.name)wrap.append(el('div','vb-mod',M('模块：','モジュール：')+x.name));
  x.r.findings.forEach(function(f){var row=el('div','vb-row vb-'+f.sev),head=el('div','vb-head');head.append(el('span','vb-badge',sevName(f.sev)),el('span','vb-line',M('第 '+f.line+' 行',f.line+' 行目')));
   row.append(head,el('div','vb-msg',ja()?f.msgJa:f.msg));if(f.fix||f.fixJa)row.append(el('div','vb-fix',M('怎么改：','直し方：')+(ja()?f.fixJa:f.fix)));
   row.title=M('点一下在代码里找到这一行','クリックするとコード内のこの行を選択します');row.onclick=function(){show(x,f.line);};wrap.append(row);});});
 if(!any)wrap.append(el('p','vb-ok',M('没有发现问题。注意：这里只检查写法，不运行代码，计算结果对不对还要实际试一下。','問題は見つかりませんでした。書き方だけのチェックで、実行はしていません。計算結果が正しいかは実際に試してください。')));
 var fixable=results.some(function(x){return x.r.findings.some(function(f){return f.autofix;});});$('vbFix').hidden=!fixable;$('vbFixed').hidden=true;$('vbCopy').hidden=true;
 say(M('检查完了：错误 '+total.error+'、安全风险 '+total.risk+'、容易出错 '+total.warn+'、建议 '+total.tip+'。','チェックしました：エラー '+total.error+'、セキュリティ '+total.risk+'、注意 '+total.warn+'、ヒント '+total.tip+'。'));}
// 在代码框里选中那一行
function show(x,line){var ta=$('vbCode');if(modules.length&&x.name&&$('vbModule').value!==x.name){$('vbModule').value=x.name;ta.value=x.code;}var lines=ta.value.split(/\r\n|\r|\n/),start=0;for(var i=0;i<line-1&&i<lines.length;i++)start+=lines[i].length+(ta.value.indexOf('\r\n')>=0?2:1);ta.focus();try{ta.setSelectionRange(start,start+(lines[line-1]||'').length);var lh=parseFloat(getComputedStyle(ta).lineHeight)||16;ta.scrollTop=Math.max(0,(line-3)*lh);}catch(ignore){}}
function fix(){var out=[],n=0;results.forEach(function(x){var r=V.autofix(x.code);n+=r.changes.length;out.push((results.length>1?"' ==== "+(x.name||'')+' ====\r\n':'')+r.code);});var ta=$('vbFixed');ta.value=out.join('\r\n\r\n');ta.hidden=false;$('vbCopy').hidden=false;
 say(M('已生成修改后的代码（改了 '+n+' 处：Integer 改成 Long、加 Option Explicit、Declare 加 PtrSafe）。复制后到 VBA 编辑器里替换原来的代码。其他问题要自己改。','修正後のコードを作成しました（'+n+' か所：Integer → Long、Option Explicit の追加、Declare に PtrSafe）。コピーして VBA エディターで元のコードと置き換えてください。その他の問題は手動で直してください。'));}
function copy(){var ta=$('vbFixed');ta.focus();ta.select();var ok=false;try{ok=document.execCommand('copy');}catch(ignore){}if(!ok&&root.navigator&&navigator.clipboard)navigator.clipboard.writeText(ta.value).then(function(){say(M('已复制。','コピーしました。'));},function(){say(M('没能自动复制，请按 Ctrl+C。','自動でコピーできませんでした。Ctrl+C を押してください。'),true);});else say(ok?M('已复制。','コピーしました。'):M('请按 Ctrl+C 复制。','Ctrl+C でコピーしてください。'));}
function task(fn){if(busy)return;busy=true;try{fn();}catch(e){say(String(e&&e.message||e),true);}busy=false;}
function build(){var host=document.querySelector('section[data-panel="files"]'),office=!host;if(office)host=$('toolSection');if(!host||!V)return;var old=$('macroCheckSection');if(old)old.remove();
 box=el('section','tool-card mask vb');box.id='macroCheckSection';box.setAttribute('translate','no');if(!office){box.dataset.panel='files';box.hidden=host.hidden;}
 var head=el('div','card-heading'),icon=el('span','section-icon teal'),img=el('img');img.src=(office?'../':'')+'icons/macrocheck.svg';img.alt='';icon.append(img);var h=el('div');h.append(el('h2','',M('宏代码检查','マクロのチェック')),el('p','',M('不运行、不用 AI，找出写错的地方、容易出 bug 的写法和有安全风险的代码','実行せず AI も使わずに、書き間違い・不具合の原因になりやすい書き方・危険なコードを見つけます')));head.append(icon,h);
 var bar=el('div','button-pair'),readB=el('button','',M('读取当前文件的宏','今のファイルのマクロを読み取る')),runB=el('button','primary',M('检查','チェック'));readB.type=runB.type='button';
 readB.onclick=function(){task(function(){modules=readMacros();var sel=$('vbModule');sel.textContent='';var all=el('option','',M('全部模块（'+modules.length+' 个）','すべてのモジュール（'+modules.length+' 個）'));all.value='__all';sel.append(all);modules.forEach(function(m){var o=el('option','',m.name);o.value=m.name;sel.append(o);});var p=el('option','',M('（用下面框里的代码）','（下の欄のコードを使う）'));p.value='__paste';sel.append(p);sel.hidden=false;$('vbCode').value=modules.map(function(m){return "' ==== "+m.name+" ====\r\n"+m.code;}).join('\r\n\r\n');run();});};
 runB.onclick=function(){task(run);};bar.append(readB,runB);
 var sel=el('select','vb-module');sel.id='vbModule';sel.hidden=true;sel.onchange=function(){var m=modules.find(function(x){return x.name===sel.value;});if(m)$('vbCode').value=m.code;else if(sel.value==='__all')$('vbCode').value=modules.map(function(m){return "' ==== "+m.name+" ====\r\n"+m.code;}).join('\r\n\r\n');task(run);};
 var ta=el('textarea','vb-code');ta.id='vbCode';ta.rows=10;ta.spellcheck=false;ta.placeholder=M('把 VBA 代码粘贴到这里（Alt+F11 打开 VBA 编辑器，全选复制）','VBA のコードをここに貼り付けてください（Alt+F11 で VBA エディターを開き、全選択してコピー）');ta.oninput=function(){if(modules.length&&$('vbModule').value!=='__paste')$('vbModule').value='__paste';};
 var res=el('div','vb-result');res.id='vbResult';var st=el('p','mk-status');st.id='vbStatus';st.setAttribute('role','status');
 var fixB=el('button','',M('生成修改后的代码','修正後のコードを作成'));fixB.type='button';fixB.id='vbFix';fixB.hidden=true;fixB.onclick=function(){task(fix);};
 var fixed=el('textarea','vb-code');fixed.id='vbFixed';fixed.rows=8;fixed.readOnly=true;fixed.hidden=true;var copyB=el('button','',M('复制','コピー'));copyB.type='button';copyB.id='vbCopy';copyB.hidden=true;copyB.onclick=copy;
 box.append(head,bar,sel,ta,st,res,fixB,fixed,copyB,el('p','mk-meta',M('只检查写法，不运行代码，也不会改你的文件。','書き方だけをチェックし、コードの実行やファイルの変更はしません。')));
 if(office)($('maskSection')||host).after(box);else{var anchor=$('maskSection')||$('mailMergeSection');if(anchor)anchor.after(box);else host.parentNode.insertBefore(box,host);}}
function start(){build();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start);else start();
try{if(root.OfficeI18n&&root.OfficeI18n.onChange)root.OfficeI18n.onChange(function(){var code=$('vbCode')&&$('vbCode').value;build();if(code)$('vbCode').value=code;if(results.length)render();});}catch(ignore){}
root.VbaPanel={readMacros:readMacros};
})(window);
