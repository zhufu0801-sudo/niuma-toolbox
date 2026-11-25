// 办公工具箱 · 数据脱敏侧栏（2026-11-25）。在侧栏“文件”页（Office 版在工具列表下面）。
// 规则：只认号码和词表里的词，不猜姓名地址；默认另存副本，原文件不动；对照表加密存在本机，可以还原。
// 界面文字直接按当前语言写（M(中文, 日文)），整块不交给整页翻译；切换语言时重画。
(function(root){'use strict';
var E=root.MaskEngine,H=root.MaskHost,box=null,S=null,busy=false,found=[],foundModules=[],lastScope='';
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
function el(tag,cls,text){var e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;}
function $(id){return document.getElementById(id);}
function ask(msg){try{if(root.OfficeTools&&root.OfficeTools.lib&&root.OfficeTools.lib.ask)return root.OfficeTools.lib.ask(msg);}catch(ignore){}var i=root.OfficeI18n;return i&&typeof i.confirm==='function'?i.confirm(msg):root.confirm(msg);}
function lang(){return ja()?'ja':'zh';}
async function post(body){var f=root.fetch.bind(root),s=await f('/api/session').then(function(r){if(!r.ok)throw Error(M('本机服务没有连接。','ローカルサービスに接続できません。'));return r.json();});
 var r=await f('/api/tools/mask',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(Object.assign({lang:lang()},body))}),d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(d.error||M('没有完成，请重试。','完了できませんでした。もう一度お試しください。'));return d;}
function say(m,bad){var s=$('mkStatus');if(!s)return;s.textContent=m||'';s.className='mk-status'+(bad?' bad':'');}
function task(fn){if(busy)return;busy=true;box.querySelectorAll('button,input,select,textarea').forEach(function(b){b.disabled=true;});
 Promise.resolve().then(fn).catch(function(e){var m=String(e&&e.message||e);try{if(ja()&&root.OfficeI18n&&root.OfficeI18n.t)m=root.OfficeI18n.t(m);}catch(ignore){}say(/^(已取消|キャンセル)/.test(m)?m:M('未完成：','完了できませんでした：')+m,true);}).then(function(){busy=false;box.querySelectorAll('button,input,select,textarea').forEach(function(b){b.disabled=false;});sync();});}
var MODES=[['star','中间打码','一部を伏せ字','138****5678、张*','090-****-5678、田*'],['block','整段遮住','全部を黒塗り','■■■■■■','■■■■■■'],['code','换成代号','記号に置き換え','【电话1】、【名称A】，同一个始终同一个代号','【電話1】、【名称A】（同じ内容は同じ記号）']];
function scope(){var r=box.querySelector('input[name="mkScope"]:checked');return r?r.value:'all';}
function mode(){var r=box.querySelector('input[name="mkMode"]:checked');return r?r.value:'star';}
function types(){return Array.prototype.slice.call(box.querySelectorAll('input[name="mkType"]:checked')).map(function(x){return x.value;});}
function words(){var p=E.parseWords($('mkWords').value);return types().indexOf('WORD')>=0?p.words:[];}
function opts(){return {types:types().filter(function(t){return t!=='WORD';}),words:words()};}
function isOffice(){return !(root.OfficeTools&&root.OfficeTools.app&&root.OfficeTools.app());}
function modeName(m){var x=MODES.find(function(v){return v[0]===m;});return x?M(x[1],x[2]):m;}
function hostName(h){return ({wps:M('WPS 文字','WPS 文字'),et:M('WPS 表格','WPS 表計算'),wpp:M('WPS 演示','WPS プレゼンテーション'),'office-word':'Word','office-excel':'Excel',modules:M('卡片 / 粘贴栏','カード / 貼り付けトレイ')})[h]||h;}
function when(t){var d=new Date(t),p=function(x){return String(x).padStart(2,'0');};return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate())+' '+p(d.getHours())+':'+p(d.getMinutes());}

// ---------- 词表方案 ----------
function profile(){if(!S)return null;return S.profiles.find(function(p){return p.id===S.current;})||S.profiles[0]||null;}
function fillProfiles(){var sel=$('mkProfile');sel.textContent='';(S.profiles.length?S.profiles:[{id:'',name:M('（还没有词表）','（単語リストなし）')}]).forEach(function(p){var o=el('option','',p.name);o.value=p.id;sel.append(o);});var cur=profile();sel.value=cur?cur.id:'';$('mkWords').value=cur?cur.words:'';$('mkProfileName').value=cur?cur.name:'';}
async function saveProfiles(extra){var list=S.profiles.slice(),cur=profile(),name=$('mkProfileName').value.trim()||M('方案 1','リスト 1'),w=$('mkWords').value;
 if(cur){cur.name=name;cur.words=w;}else if(w.trim()){list.push({name:name,words:w});}
 var parsed=E.parseWords(w);var d=await post({action:'profiles',profiles:list,current:cur?cur.id:'',prefs:prefs()});S.profiles=d.profiles;S.current=d.current;S.prefs=d.prefs;fillProfiles();if(!extra)say(M('词表已保存（加密）：','単語リストを保存しました（暗号化）：')+parsed.words.length+M(' 个词。',' 語。')+(parsed.bad.length?M('有 '+parsed.bad.length+' 行太长或格式不对，没收进去。',parsed.bad.length+' 行は長すぎるか形式が正しくないため除外しました。'):''));}
function prefs(){return {mode:mode(),ai:$('mkAi').checked,agent:$('mkAgent').checked,autoModules:$('mkAuto').checked,types:types()};}
async function savePrefs(){if(!S)return;try{S.prefs=(await post({action:'prefs',prefs:prefs()})).prefs;}catch(e){say(e.message,true);}}

// ---------- 找一找 ----------
function preview(list){var sess=E.session({mode:mode(),lang:lang(),words:words()});return list.map(function(x){return sess.replace(x.value,x.type);});}
function renderFound(){var wrap=$('mkFound');wrap.textContent='';if(!found.length&&!foundModules.length){wrap.hidden=true;return;}wrap.hidden=false;
 var head=el('div','mk-found-head'),all=el('button','link',M('全选','すべて選択')),none=el('button','link',M('全不选','選択解除'));all.type=none.type='button';all.onclick=function(){wrap.querySelectorAll('input[type=checkbox]').forEach(function(c){c.checked=true;});};none.onclick=function(){wrap.querySelectorAll('input[type=checkbox]').forEach(function(c){c.checked=false;});};
 if(foundModules.length){head.append(el('span','',M('找到 '+foundModules.length+' 条有要脱敏的内容','マスキング対象を含むもの：'+foundModules.length+' 件')),all,none);wrap.append(head);
  foundModules.forEach(function(x,i){var l=el('label','mk-row'),c=el('input');c.type='checkbox';c.checked=true;c.dataset.i=i;var t=el('span','mk-val',(x.kind==='card'?M('卡片：','カード：'):M('粘贴栏：','トレイ：'))+x.title);var m=el('span','mk-meta',x.count+M(' 处',' か所')+(x.linked?M('（联动模块：以后“更新联动内容”会重新拿回原文）','（連動モジュール：「連動内容を更新」すると元の内容に戻ります）'):''));l.append(c,t,m);wrap.append(l);});return;}
 var total=found.reduce(function(s,x){return s+x.count;},0);head.append(el('span','',M('找到 '+found.length+' 种、共 '+total+' 处','見つかった内容：'+found.length+' 種類・計 '+total+' か所')),all,none);wrap.append(head);
 var pv=preview(found);found.forEach(function(x,i){var l=el('label','mk-row'),c=el('input');c.type='checkbox';c.checked=x.on!==false;c.dataset.i=i;c.onchange=function(){x.on=c.checked;};var v=el('span','mk-val',x.value),arrow=el('span','mk-to','→ '+pv[i]),m=el('span','mk-meta',E.label(x.type,ja())+' · '+x.count+M(' 处',' か所'));l.append(c,v,arrow,m);wrap.append(l);});}
async function scan(){foundModules=[];found=[];var sc=scope();lastScope=sc;if(!types().length)throw Error(M('请至少勾一种要认的内容。','対象の種類を 1 つ以上選んでください。'));if(types().length===1&&types()[0]==='WORD'&&!words().length)throw Error(M('词表是空的：请在下面填上要隐藏的词（客户名、项目名等），一行一个。','単語リストが空です。隠したい語（顧客名・案件名など）を 1 行に 1 つ入力してください。'));
 if(sc==='modules'){say(M('正在检查卡片和粘贴栏…','カードと貼り付けトレイを確認しています…'));foundModules=(await post({action:'modules-scan',types:types(),words:$('mkWords').value})).items;renderFound();say(foundModules.length?M('勾选要脱敏的，再点“脱敏”。','マスキングするものを選んで「マスキング」を押してください。'):M('卡片和粘贴栏里没有找到要脱敏的内容。','カードと貼り付けトレイにマスキング対象はありませんでした。'));return;}
 say(M('正在检查…','確認しています…'));var ad=H.adapter(),segs=await ad.segments(sc),o=opts(),agg={},order=[];
 segs.forEach(function(seg){E.find(seg.text,o).forEach(function(h){if(!ad.inside(seg,h))return;var k=h.type+'\u0001'+h.value;if(!agg[k]){agg[k]={value:h.value,type:h.type,count:0};order.push(agg[k]);}agg[k].count++;});});
 var rank={};E.TYPES.forEach(function(t,i){rank[t.type]=i;});found=order.sort(function(a,b){return rank[a.type]-rank[b.type]||b.count-a.count;});renderFound();
 var extra='';try{var nm=ad.names?ad.names(words()):[];if(nm.length)extra=M('\n工作表名称里也有词表的词：'+nm.join('、')+'。名称不会自动改，请自己改一下。','\nシート名にも単語リストの語があります：'+nm.join('、')+'。シート名は自動で変えないので、手動で直してください。');}catch(ignore){}
 say((found.length?M('检查完了。看一下清单，不想换的取消勾选，再点“脱敏”。','確認しました。一覧を見て、変えたくないものはチェックを外し「マスキング」を押してください。'):M('没有找到要脱敏的内容。','マスキング対象は見つかりませんでした。'))+extra+(sc==='all'&&!isOffice()?M('\n图片和截图里的文字认不出来，请自己看一下。','\n画像やスクリーンショット内の文字は認識できません。目視で確認してください。'):''));}

// ---------- 脱敏 ----------
function chosen(){var set={};$('mkFound').querySelectorAll('input[type=checkbox]').forEach(function(c){if(c.checked){var x=found[Number(c.dataset.i)];if(x)set[x.type+'\u0001'+x.value]=1;}});return set;}
async function apply(){var sc=scope();if(sc!==lastScope||(!found.length&&!foundModules.length))await scan();if(sc==='modules')return applyModules();if(!found.length)return;
 var pick=chosen(),kinds=Object.keys(pick).length;if(!kinds)throw Error(M('清单里一个都没勾。','一覧で何も選ばれていません。'));
 var ad=H.adapter(),info=await ad.info(),office=!!ad.office,inplace=office||$('mkInplace').checked,m=mode(),total=found.filter(function(x){return pick[x.type+'\u0001'+x.value];}).reduce(function(s,x){return s+x.count;},0);
 var target=null;if(!inplace)target=await post({action:'target',source:info.path,name:info.name,suffix:'mask'});
 var msg=M('将把 '+kinds+' 种内容（共 '+total+' 处）'+modeName(m)+'。\n','選んだ '+kinds+' 種類（計 '+total+' か所）を「'+modeName(m)+'」にします。\n')+(inplace?(office?M('Office 版不能自动另存副本，会直接改当前文件。还没另存过副本的话，请先点“取消”，用「文件 → 另存为」存一份副本，在副本里再做。\n','Office 版では自動でコピーを保存できないため、今のファイルを直接変更します。まだコピーを保存していない場合は「キャンセル」し、「ファイル → 名前を付けて保存」でコピーを作ってから実行してください。\n'):M('直接改当前文件（不另存副本）。\n','今のファイルを直接変更します（コピーを作りません）。\n')):M('先另存为副本：\n'+target.path+'\n在副本里改，原文件不动。\n','先にコピーを保存し、コピーの方を変更します（元のファイルはそのまま）：\n'+target.path+'\n'))+M('对照表加密保存在这台电脑上，以后可以还原。是否继续？','対応表はこのパソコンに暗号化して保存し、あとで元に戻せます。続行しますか？');
 if(!ask(msg))throw Error(M('已取消。','キャンセルしました。'));
 if(!inplace){if(info.hasPath&&!info.saved&&ask(M('原文件有还没保存的修改。先保存原文件吗？\n“确定”：保存后再做副本；“取消”：不保存，这些修改只进副本。','元のファイルに未保存の変更があります。先に保存しますか？\n「OK」：保存してからコピー／「キャンセル」：保存せず、変更はコピーにだけ入ります。'))){try{(root.OfficeTools.host()==='et'?root.OfficeTools.app().ActiveWorkbook:root.OfficeTools.host()==='wpp'?root.OfficeTools.app().ActivePresentation:root.OfficeTools.app().ActiveDocument).Save();}catch(e){throw Error(M('原文件没能保存：','元のファイルを保存できませんでした：')+e.message);}}
  say(M('正在另存副本…','コピーを保存しています…'));ad.copy(target.path);}
 var note=await ad.prepare(!inplace);say(M('正在脱敏…','マスキングしています…'));
 var segs=await ad.segments(sc),o=opts(),sess=E.session({mode:m,lang:lang(),words:words()}),data={scope:sc,exact:ad.exact?{}:null,seq:ad.exact?null:{}},counts={},all=[],skipped=0;
 function keep(h){return pick[h.type+'\u0001'+h.value];}
 function count(p){all.push(p);counts[p[2]]=(counts[p[2]]||0)+1;}
 if(ad.exact){var items=[];segs.forEach(function(seg){var hits=E.find(seg.text,o).filter(function(h){return keep(h)&&ad.inside(seg,h);});if(!hits.length)return;var r=E.apply(seg.text,hits,sess);items.push({seg:seg,value:r.text,seq:r.seq});});
  var entries=ad.putMany?await ad.putMany(items):items.map(function(x){try{return ad.put(x.seg,x.value);}catch(e){return null;}});
  items.forEach(function(x,i){if(!entries[i]){skipped+=x.seq.length;return;}data.exact[x.seg.key]=Object.assign({k:x.seg.kind||'cell'},entries[i]);x.seq.forEach(count);});}
 else{var jobs=[];segs.forEach(function(seg){var hits=E.find(seg.text,o).filter(function(h){return keep(h)&&ad.inside(seg,h);});if(!hits.length)return;var r=E.apply(seg.text,hits,sess);jobs.push({seg:seg,seq:r.seq,edits:hits.map(function(h,i){return {start:h.start,end:h.end,text:r.seq[i][0],expect:h.value};})});});
  var oks=ad.writeMany?await ad.writeMany(jobs):jobs.map(function(j){return ad.write(j.seg,j.edits);});
  jobs.forEach(function(j,i){j.seq.forEach(function(p,k){if(oks[i]&&oks[i][k]){(data.seq[j.seg.key]=data.seq[j.seg.key]||[]).push(p);count(p);}else skipped++;});});}
 var cleaned=await ad.clean(),after=await ad.info();
 if(!all.length){if(!inplace)try{ad.save();}catch(ignore){}throw Error(M('一处也没能改（位置对不上）。','1 か所も変更できませんでした（位置が合いません）。'));}
 var rec=null;try{rec=(await post({action:'save',record:{kind:'doc',host:ad.host,name:after.name,path:after.path,source:inplace?'':info.path,mode:m,counts:counts,total:all.length,pairs:E.table([all]),data:data}})).record;}
 catch(e){if(!inplace)try{ad.save();}catch(ignore){}throw Error(M('已经脱敏，但对照表没保存成功（'+e.message+'），这份文件不能自动还原。','マスキングしましたが、対応表を保存できませんでした（'+e.message+'）。このファイルは自動で元に戻せません。'));}
 await ad.mark(rec.id);if(!inplace)ad.save();else if(ad.undoClear)ad.undoClear();
 found=[];renderFound();await refresh();
 say(M('已脱敏 '+all.length+' 处。','マスキングしました：'+all.length+' か所。')+(inplace?(office?M('\n记得保存文件。','\nファイルを保存してください。'):M('\n改的是当前文件，记得保存。','\n今のファイルを変更しました。保存してください。')):M('\n副本已保存：'+after.path+'\n原文件没有改动。','\nコピーを保存しました：'+after.path+'\n元のファイルは変更していません。'))+
  M('\n对照表已加密保存，点下面“还原”可以换回原文。','\n対応表を暗号化して保存しました。下の「元に戻す」で元の内容に戻せます。')+(skipped?M('\n有 '+skipped+' 处因为文字里有域或隐藏内容、位置对不上，没有改，请手动看一下。','\n'+skipped+' か所はフィールドや隠し文字があり位置が合わないため、変更していません。手動で確認してください。'):'')+(note?'\n'+note:'')+(cleaned?M('\n文件属性里的作者、单位已清空。','\nファイルのプロパティの作成者・会社名を消去しました。'):'')+
  (office?M('\nOffice 版只处理了'+(ad.host==='office-word'?'正文（页眉页脚、文本框、批注请手动检查）':'单元格（批注、页眉页脚请手动检查）')+'。','\nOffice 版では'+(ad.host==='office-word'?'本文だけ（ヘッダー・フッター、テキストボックス、コメントは手動で確認）':'セルだけ（コメント、ヘッダー・フッターは手動で確認）')+'を処理しました。'):''));}
async function applyModules(){var pick=[];$('mkFound').querySelectorAll('input[type=checkbox]').forEach(function(c){if(c.checked){var x=foundModules[Number(c.dataset.i)];if(x)pick.push({kind:x.kind,id:x.id});}});if(!pick.length)throw Error(M('清单里一个都没勾。','一覧で何も選ばれていません。'));
 if(!ask(M('将把选中的 '+pick.length+' 条卡片 / 粘贴栏内容'+modeName(mode())+'（直接改这几条）。对照表加密保存，以后可以还原。是否继续？','選んだ '+pick.length+' 件のカード / 貼り付けトレイの内容を「'+modeName(mode())+'」にします（そのまま変更します）。対応表は暗号化して保存し、あとで元に戻せます。続行しますか？')))throw Error(M('已取消。','キャンセルしました。'));
 var d=await post({action:'modules-apply',items:pick,mode:mode(),types:types(),words:$('mkWords').value});foundModules=[];renderFound();await refresh();say(d.changed?M('已脱敏 '+d.changed+' 条、共 '+d.total+' 处。可以在下面的对照表里还原。','マスキングしました：'+d.changed+' 件・計 '+d.total+' か所。下の対応表から元に戻せます。'):M('没有需要改的内容。','変更する内容はありませんでした。'));}

// ---------- 还原 ----------
function groupSegs(segs,key){if(/#\*$/.test(key)){var pre=key.slice(0,-1);return segs.filter(function(s){return s.key.indexOf(pre)===0;});}return segs.filter(function(s){return s.key===key;});}
// 一组片段按顺序拼起来找：每处脱敏结果按记下的顺序依次往后找，找到就换回原文；找不到的（后来被改过）跳过。
function restorePlan(group,seq){var joined='',offs=[];group.forEach(function(s){offs.push(joined.length);joined+=s.text+'\u0001';});var pos=0,edits=group.map(function(){return [];}),missing=0;
 seq.forEach(function(p){var i=joined.indexOf(p[0],pos);if(i<0){missing++;return;}var g=0;while(g+1<offs.length&&offs[g+1]<=i)g++;edits[g].push({start:i-offs[g],end:i-offs[g]+p[0].length,text:p[1],expect:p[0]});pos=i+p[0].length;});return {edits:edits,missing:missing};}
async function restoreDoc(id){var ad=H.adapter(),info=await ad.info();
 if(!id){id=await ad.marked();if(!id&&info.path){var hit=(await post({action:'find',path:info.path})).records;if(hit.length)id=hit[0].id;}}
 if(!id)throw Error(M('这个文件里没有找到脱敏记录。请在下面“对照表”里找到它，点那一行的“还原到当前文件”。','このファイルにはマスキングの記録が見つかりません。下の「対応表」から該当する行を探し、「今のファイルに戻す」を押してください。'));
 var rec=await post({action:'get',id:id});if(rec.kind!=='doc')throw Error(M('这份对照表是卡片 / 粘贴栏的，请点那一行的“还原”。','この対応表はカード / 貼り付けトレイ用です。その行の「元に戻す」を押してください。'));
 if(rec.host!==ad.host)throw Error(M('这份对照表是在'+hostName(rec.host)+'里做的，请在'+hostName(rec.host)+'里打开文件再还原。','この対応表は '+hostName(rec.host)+' で作られました。'+hostName(rec.host)+' でファイルを開いてから元に戻してください。'));
 var office=!!ad.office,inplace=office||$('mkInplace').checked,target=inplace?null:await post({action:'target',source:info.path,name:info.name,suffix:'restore'});
 if(!ask(M('将按对照表（'+when(rec.created)+'，'+rec.total+' 处）把脱敏的内容换回原文。\n','対応表（'+when(rec.created)+'・'+rec.total+' か所）に従って、マスキングした内容を元に戻します。\n')+(inplace?M('直接改当前文件。','今のファイルを直接変更します。'):M('另存为：\n'+target.path+'\n脱敏的文件不动。','次の名前で保存します（マスキングしたファイルはそのまま）：\n'+target.path))+M('\n是否继续？','\n続行しますか？')))throw Error(M('已取消。','キャンセルしました。'));
 if(!inplace)ad.copy(target.path);await Promise.resolve(ad.prepare(!inplace));
 var data=rec.data||{},segs=await ad.segments('all'),done=0,missing=0;
 if(data.exact){var map={};segs.forEach(function(s){map[s.key]=s;});var keys=Object.keys(data.exact);
  if(ad.backMany){var oks=await ad.backMany(keys.map(function(k){return {key:k,entry:data.exact[k]};}));oks.forEach(function(ok){if(ok)done++;else missing++;});}
  else keys.forEach(function(k){var s=map[k];if(!s){missing++;return;}var ok=false;try{ok=ad.back(s,data.exact[k]);}catch(e){ok=false;}if(ok)done++;else missing++;});}
 else{var jobs=[];Object.keys(data.seq||{}).forEach(function(k){var group=groupSegs(segs,k),plan=restorePlan(group,data.seq[k]);missing+=plan.missing;group.forEach(function(s,i){if(plan.edits[i].length)jobs.push({seg:s,edits:plan.edits[i]});});});
  var res=ad.writeMany?await ad.writeMany(jobs):jobs.map(function(j){return ad.write(j.seg,j.edits);});res.forEach(function(r){r.forEach(function(ok){if(ok)done++;else missing++;});});}
 var after=await ad.info();if(!inplace)ad.save();await post({action:'restored',id:id,path:after.path});await refresh();
 say(M('已还原 '+done+' 处。','元に戻しました：'+done+' か所。')+(missing?M('\n有 '+missing+' 处没找到（脱敏后文字被改过或删掉了），请对照“查看对照表”手动补上。','\n'+missing+' か所は見つかりませんでした（マスキング後に文字が変更・削除された可能性があります）。「対応表を見る」を参照して手動で直してください。'):'')+(inplace?M('\n记得保存文件。','\nファイルを保存してください。'):M('\n已另存为：'+after.path,'\n保存しました：'+after.path))+M('\n作者、单位等文件属性脱敏时已清空，不会还原。','\n作成者・会社名などのプロパティはマスキング時に消去したため、元に戻りません。'));}

// ---------- 对照表列表 ----------
async function refresh(){S=await post({action:'state'});renderRecords();}
function renderRecords(){var list=$('mkRecords');list.textContent='';$('mkProtect').hidden=!S||S.protection==='dpapi'||S.protection==='keychain';
 if(!S||!S.records.length){list.append(el('p','mk-empty',M('还没有对照表。','対応表はまだありません。')));return;}
 S.records.slice(0,30).forEach(function(r){var nm=r.kind==='clip'?M('复制给 AI','AI 用にコピー'):r.kind==='modules'?(/自动/.test(r.name)?M('卡片 / 粘贴栏（自动脱敏）','カード / 貼り付けトレイ（自動マスキング）'):M(r.name,String(r.name).replace(/卡片 \/ 粘贴栏 (\d+) 条/,'カード / 貼り付けトレイ $1 件'))):r.name;var row=el('div','mk-rec'),title=el('div','mk-rec-title',nm||M('（未命名）','（名前なし）')),meta=el('div','mk-meta',when(r.created)+' · '+hostName(r.host)+' · '+modeName(r.mode)+' · '+r.total+M(' 处',' か所')+(r.restored?M(' · 已还原过',' · 復元済み'):''));
  var bar=el('div','mk-rec-bar'),view=el('button','link',M('查看','見る')),exp=el('button','link',M('导出','書き出す')),back=el('button','link',r.kind==='modules'?M('还原','元に戻す'):M('还原到当前文件','今のファイルに戻す')),del=el('button','link',M('删除','削除'));[view,exp,back,del].forEach(function(b){b.type='button';});
  var table=el('div','mk-table');table.hidden=true;
  view.onclick=function(){if(!table.hidden){table.hidden=true;return;}task(async function(){var d=await post({action:'get',id:r.id});table.textContent='';var t=el('table');var h=el('tr');[M('原文','元の内容'),M('替换为','置き換え後'),M('类型','種類'),M('次数','回数')].forEach(function(x){h.append(el('th','',x));});t.append(h);(d.pairs||[]).slice(0,500).forEach(function(p){var tr=el('tr');[p.original,p.masked,E.label(p.type,ja()),String(p.count)].forEach(function(x){tr.append(el('td','',x));});t.append(tr);});table.append(t);if((d.pairs||[]).length>500)table.append(el('p','mk-meta',M('只显示前 500 行，全部内容请导出。','先頭 500 行だけ表示しています。すべて見るには書き出してください。')));table.hidden=false;say('');});};
  exp.onclick=function(){if(!ask(M('导出的表格里是原文（没有加密）。用完请删掉，不要随文件一起发出去。是否导出？','書き出した表には元の内容がそのまま入ります（暗号化されません）。使い終わったら削除し、ファイルと一緒に送らないでください。書き出しますか？')))return;task(async function(){var d=await post({action:'export',id:r.id,open:true});say(M('已导出：','書き出しました：')+d.path);});};
  back.onclick=function(){task(async function(){if(r.kind==='modules'){if(!ask(M('把这些卡片 / 粘贴栏内容换回原文吗？','これらのカード / 貼り付けトレイの内容を元に戻しますか？')))throw Error(M('已取消。','キャンセルしました。'));var d=await post({action:'modules-restore',id:r.id});await refresh();say(M('已还原 '+d.restored+' 条。','元に戻しました：'+d.restored+' 件。')+(d.skipped.length?M('\n没还原：','\n戻せなかったもの：')+d.skipped.join('、'):''));return;}await restoreDoc(r.id);});};
  del.onclick=function(){if(!ask(M('删除后这份文件就不能还原了。确定删除这份对照表吗？','削除すると、このファイルは元に戻せなくなります。この対応表を削除しますか？')))return;task(async function(){await post({action:'delete',id:r.id});await refresh();say(M('已删除。','削除しました。'));});};
  bar.append(view,exp);if(r.kind!=='clip')bar.append(back);bar.append(del);row.append(title,meta,bar,table);list.append(row);});}

// ---------- 界面 ----------
function sync(){if(!box)return;var sc=scope(),office=isOffice();$('mkInplaceRow').hidden=sc==='modules'||office;$('mkOfficeNote').hidden=!office||sc==='modules';var go=$('mkApply');go.textContent=sc==='modules'?M('脱敏','マスキング'):office||$('mkInplace').checked?M('脱敏（改当前文件）','マスキング（今のファイル）'):M('脱敏并另存副本','マスキングしてコピーを保存');}
function build(){var host=document.querySelector('section[data-panel="files"]'),office=!host;if(office)host=$('toolSection');if(!host||!E||!H)return;
 var old=$('maskSection');if(old)old.remove();box=el('section','tool-card mask');box.id='maskSection';box.setAttribute('translate','no');if(!office){box.dataset.panel='files';box.hidden=host.hidden;}
 var head=el('div','card-heading'),icon=el('span','section-icon teal'),img=el('img');img.src=(office?'../':'')+'icons/shield.svg';img.alt='';icon.append(img);var h=el('div');h.append(el('h2','',M('数据脱敏','データのマスキング')),el('p','',M('号码和指定的词换成 * 或代号；默认另存副本，原文件不动，可以还原','番号や指定した語を伏せ字・記号に置き換え。元のファイルはそのまま、あとで元に戻せます')));head.append(icon,h);
 function radios(name,list,val){var fs=el('fieldset','mk-group');list.forEach(function(x){var l=el('label'),r=el('input');r.type='radio';r.name=name;r.value=x[0];r.checked=x[0]===val;r.onchange=function(){found=[];foundModules=[];renderFound();sync();if(name==='mkMode')savePrefs();};l.append(r,el('span','',x[1]));if(x[2]){l.append(el('small','mk-ex',x[2]));}fs.append(l);});return fs;}
 var scopes=radios('mkScope',[['all',M('整个文件（含页眉页脚、批注、文本框、备注）','ファイル全体（ヘッダー・フッター、コメント、テキストボックス、ノートを含む）')],['selection',M('选中的内容 / 单元格','選択した内容 / セル')],['modules',M('卡片和粘贴栏','カードと貼り付けトレイ')]],'all');scopes.prepend(el('legend','',M('处理哪里','対象')));
 var tf=el('fieldset','mk-group mk-types');tf.append(el('legend','',M('认哪些（不猜姓名、地址）','対象の種類（氏名・住所の推測はしません）')));E.TYPES.forEach(function(t){var l=el('label'),c=el('input');c.type='checkbox';c.name='mkType';c.value=t.type;c.checked=!S||S.prefs.types.indexOf(t.type)>=0;c.onchange=function(){found=[];renderFound();savePrefs();};l.append(c,el('span','',ja()?t.labelJa:t.label));tf.append(l);});
 var wf=el('fieldset','mk-group');wf.append(el('legend','',M('词表（客户名、项目名、人名等）','単語リスト（顧客名・案件名・人名など）')));var prow=el('div','mk-prow'),sel=el('select');sel.id='mkProfile';sel.onchange=function(){S.current=sel.value;fillProfiles();found=[];renderFound();};var pname=el('input');pname.id='mkProfileName';pname.placeholder=M('方案名','リスト名');var add=el('button','',M('新建','新規')),rm=el('button','',M('删除','削除'));add.type=rm.type='button';
 add.onclick=function(){task(async function(){var list=S.profiles.slice();list.push({name:M('方案 ','リスト ')+(list.length+1),words:''});var d=await post({action:'profiles',profiles:list,current:'',prefs:prefs()});S.profiles=d.profiles;S.current=d.profiles[d.profiles.length-1].id;await post({action:'profiles',profiles:S.profiles,current:S.current});fillProfiles();say(M('已新建方案，填好词后点“保存词表”。','新しいリストを作りました。語を入力して「単語リストを保存」を押してください。'));});};
 rm.onclick=function(){var cur=profile();if(!cur)return;if(!ask(M('删除词表“'+cur.name+'”吗？','単語リスト「'+cur.name+'」を削除しますか？')))return;task(async function(){var d=await post({action:'profiles',profiles:S.profiles.filter(function(p){return p.id!==cur.id;}),current:''});S.profiles=d.profiles;S.current=d.current;fillProfiles();say(M('已删除。','削除しました。'));});};
 prow.append(sel,pname,add,rm);var ta=el('textarea');ta.id='mkWords';ta.rows=5;ta.placeholder=M('一行一个，例如：\n北京某某科技有限公司\n张三=员工A（“=”后面是自己定的代号）\n# 开头的行是备注','1 行に 1 つ。例：\n株式会社サンプル\n田中太郎=社員A（「=」の後は自分で決めた記号）\n# で始まる行はメモ');ta.oninput=function(){found=[];renderFound();};
 var save=el('button','',M('保存词表','単語リストを保存'));save.type='button';save.onclick=function(){task(function(){return saveProfiles();});};
 var ai=el('label','mk-check'),aic=el('input');aic.type='checkbox';aic.id='mkAi';aic.checked=!S||S.prefs.ai;aic.onchange=savePrefs;ai.append(aic,el('span','',M('用 AI 时，发送前也隐藏这些词','AI に送る前にもこれらの語を隠す')));
 var agent=el('label','mk-check'),agentc=el('input');agentc.type='checkbox';agentc.id='mkAgent';agentc.checked=!S||S.prefs.agent!==false;agentc.onchange=savePrefs;agent.append(agentc,el('span','',M('agent 通过工具箱读文档时自动脱敏','agent がツールボックスで文書を読むとき自動でマスキングする')));
 wf.append(prow,ta,save,ai,agent);
 var modes=radios('mkMode',MODES.map(function(x){return [x[0],M(x[1],x[2]),ja()?x[4]:x[3]];}),S?S.prefs.mode:'star');modes.prepend(el('legend','',M('换成什么','置き換え方')));
 var inRow=el('label','mk-check');inRow.id='mkInplaceRow';var inc=el('input');inc.type='checkbox';inc.id='mkInplace';inc.onchange=sync;inRow.append(inc,el('span','',M('直接改当前文件（不另存副本，不推荐）','今のファイルを直接変更する（コピーを作らない・非推奨）')));
 var offNote=el('p','mk-note',M('Office 版不能自动另存副本：请先用「文件 → 另存为」存一份副本，在副本里脱敏。PowerPoint 请用 WPS 演示打开再脱敏。','Office 版では自動でコピーを保存できません。先に「ファイル → 名前を付けて保存」でコピーを作り、コピーでマスキングしてください。PowerPoint は WPS プレゼンテーションで開いてマスキングしてください。'));offNote.id='mkOfficeNote';
 var auto=el('label','mk-check'),ac=el('input');ac.type='checkbox';ac.id='mkAuto';ac.checked=!!(S&&S.prefs.autoModules);ac.onchange=savePrefs;auto.append(ac,el('span','',M('存进卡片和粘贴栏时自动脱敏（按当前词表和方式）','カードや貼り付けトレイに入れるとき自動でマスキング（今の単語リストと置き換え方で）')));
 var bar=el('div','button-pair'),scanB=el('button','',M('找一找','探す')),go=el('button','primary',M('脱敏并另存副本','マスキングしてコピーを保存'));scanB.type=go.type='button';go.id='mkApply';scanB.onclick=function(){task(scan);};go.onclick=function(){task(apply);};bar.append(scanB,go);
 var foundBox=el('div','mk-found');foundBox.id='mkFound';foundBox.hidden=true;var status=el('p','mk-status');status.id='mkStatus';status.setAttribute('role','status');
 var rest=el('details','mk-restore');rest.append(el('summary','',M('还原与对照表','元に戻す・対応表')));var rb=el('button','',M('还原当前打开的文件','今開いているファイルを元に戻す'));rb.type='button';rb.onclick=function(){task(function(){return restoreDoc('');});};
 var prot=el('p','mk-note',M('这台电脑上用不了系统加密，对照表没有加密。','このパソコンではシステムの暗号化が使えないため、対応表は暗号化されていません。'));prot.id='mkProtect';prot.hidden=true;
 var recs=el('div','mk-records');recs.id='mkRecords';rest.append(rb,el('p','mk-meta',M('对照表加密存在这台电脑上，只有你的 Windows / Mac 用户能打开；删除后就不能还原。','対応表はこのパソコンに暗号化して保存され、あなたの Windows / Mac ユーザーだけが開けます。削除すると元に戻せません。')),prot,recs);
 box.append(head,scopes,tf,wf,modes,inRow,offNote,auto,bar,foundBox,status);
 if(root.MaskClipboard)root.MaskClipboard.mount(box,{M:M,el:el,post:post,say:say,task:task,opts:opts,lang:lang,refresh:refresh});box.append(rest);
 if(office)host.after(box);else{var anchor=$('mailMergeSection')||$('pasteBarSection');if(anchor)anchor.after(box);else host.parentNode.insertBefore(box,host);}
 sync();if(S){fillProfiles();renderRecords();}}
async function start(){build();try{await refresh();build();}catch(e){if(box)say(e.message,true);}}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start);else start();
try{if(root.OfficeI18n&&root.OfficeI18n.onChange)root.OfficeI18n.onChange(function(){found=[];foundModules=[];build();});}catch(ignore){}
root.MaskPanel={scan:function(){return task(scan);},refresh:refresh,restorePlan:restorePlan,groupSegs:groupSegs};
})(window);
