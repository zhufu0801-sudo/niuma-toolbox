// 2026-11-25 数据脱敏：识别、替换、加密对照表、卡片 / 粘贴栏、AI 共用词表、WPS 三个程序的改写和还原。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import fsp from 'node:fs/promises';import os from 'node:os';import path from 'node:path';import vm from 'node:vm';
import {createMaskVault,maskModules,restoreModules,scanModules,exportRecordCsv,maskEngine} from '../lib/masking.mjs';
import {createCardStore} from '../lib/cards.mjs';import {createPasteBar} from '../lib/pastebar.mjs';
import {createPrivacyGuard} from '../lib/privacy.mjs';
import {createCoachServer} from '../server.mjs';

const same=(a,b)=>assert.equal(JSON.stringify(a),JSON.stringify(b));
const read=f=>fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8');
const fakeSecrets={available:true,async protect(p){return {value:'K'+Buffer.from(p).toString('hex'),protection:'dpapi'};},async unprotect(v){return Buffer.from(v.slice(1),'hex').toString();}};
async function tmp(t){const d=await fsp.mkdtemp(path.join(os.tmpdir(),'mask-'));t.after(()=>fsp.rm(d,{recursive:true,force:true}));return d;}
function idNo(b){const w=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2],c='10X98765432';let s=0;for(let i=0;i<17;i++)s+=Number(b[i])*w[i];return b+c[s%11];}
const ID=idNo('11010119900307123');
// 侧栏的 MaskHost / MaskPanel 在没有界面的环境里加载（只用到它们的纯函数和程序适配）
function browser(){const ctx={console};ctx.window=ctx;ctx.document={readyState:'complete',querySelector:()=>null,getElementById:()=>null};vm.runInNewContext(read('mask-engine.js')+'\n'+read('mask-host.js')+'\n'+read('mask-panel.js'),ctx);return ctx;}

test('脱敏侧栏切换文档清理命中清单和结果，不清空用户词表',()=>{
 const elements={},get=id=>elements[id]||(elements[id]={textContent:'旧结果',value:'张伟',checked:false});
 const app={ActiveDocument:{FullName:'C:/a.docx'}},ctx={OfficeTools:{app:()=>app,host:()=> 'wps'},document:{readyState:'loading',addEventListener(){},getElementById:get}};ctx.window=ctx;
 // Expose closure state only in the test VM; production keeps it private.
 const hook="root.qa={check:checkDocument,copy:copyDocument,seed:function(){box={querySelector:()=>({value:'all'})};found=[{value:'张伟'}];foundModules=[{}];lastDocument=documentKey();},counts:function(){return [found.length,foundModules.length];}};";
 vm.runInNewContext(read('mask-panel.js').replace('root.MaskPanel={',hook+'root.MaskPanel={'),ctx);ctx.qa.seed();
 ctx.qa.copy({copy:p=>{app.ActiveDocument.FullName=p;}},'C:/a-copy.docx');ctx.qa.check();same(ctx.qa.counts(),[1,1]);assert.equal(get('mkStatus').textContent,'旧结果');
 app.ActiveDocument={FullName:'C:/b.docx'};ctx.qa.check();same(ctx.qa.counts(),[0,0]);assert.equal(get('mkStatus').textContent,'');assert.equal(get('mkFound').hidden,true);assert.equal(get('mkWords').value,'张伟');
});

test('识别：词表优先、长词优先，号码按校验位认，不猜姓名',async()=>{
  const E=await maskEngine(),w=E.parseWords('北京\n北京某某科技有限公司=客户甲\n# 备注\n张三');
  same(w.words.map(x=>x.word),['北京某某科技有限公司','北京','张三']);
  const hits=E.find(`张三在北京某某科技有限公司，身份证${ID}，电话 138-1234-5678，邮箱 zs@example.com，李四没在词表里`,{words:w.words});
  same(hits.map(h=>[h.type,h.value]),[['WORD','张三'],['WORD','北京某某科技有限公司'],['ID',ID],['PHONE','138-1234-5678'],['EMAIL','zs@example.com']]);
  assert.equal(E.find('订单号 110101199003071234 校验位不对').length,0);
  assert.equal(E.find('马上 13812345678',{types:['EMAIL']}).length,0);
  assert.throws(()=>E.parseWords(Array.from({length:501},(_,i)=>'词'+i).join('\n')),/500/);
});
test('三种替换方式：打码保留首尾、遮住同长度、代号同一个原文同一个代号',async()=>{
  const E=await maskEngine();
  assert.equal(E.star('13812345678','PHONE'),'138****5678');assert.equal(E.star('+86 138 1234 5678','PHONE'),'+86 138 **** 5678');assert.equal(E.star('090-1234-5678','PHONE'),'090-****-5678');
  assert.equal(E.star(ID,'ID'),ID.slice(0,3)+'*'.repeat(11)+ID.slice(-4));assert.equal(E.star('4111 1111 1111 1111','CARD'),'**** **** **** 1111');assert.equal(E.star('zs@example.com','EMAIL'),'z***@example.com');assert.equal(E.star('张三丰','WORD'),'张**');
  assert.equal(E.block('张三'),'■■');
  const s=E.session({mode:'code',lang:'zh',words:E.parseWords('张三\n北京某某科技有限公司=客户甲').words});
  const r=E.maskText('张三给13812345678打电话，张三在北京某某科技有限公司',s,{words:E.parseWords('张三\n北京某某科技有限公司=客户甲').words});
  assert.equal(r.text,'【名称A】给【电话1】打电话，【名称A】在客户甲');
  const back=E.restore(r.text,r.seq);assert.equal(back.text,'张三给13812345678打电话，张三在北京某某科技有限公司');assert.equal(back.done,4);
  const j=E.session({mode:'code',lang:'ja'});assert.equal(j.replace('090-1111-2222','PHONE'),'【電話1】');
  // 遮住后多个原文长得一样：按顺序对回去；文字被改过的地方不乱换
  const b=E.restore('■■和■■',[['■■','张三'],['■■','李四']]);assert.equal(b.text,'张三和李四');
  const miss=E.restore('■■',[['■■','张三'],['■■','李四']]);assert.equal(miss.text,'张三');assert.equal(miss.missing.length,1);
});
test('对照表：加密保存（文件里没有原文），换电脑打不开，能导出 CSV',async t=>{
  const d=await tmp(t),f=path.join(d,'masking.json'),v=createMaskVault(f,{secrets:fakeSecrets});
  await v.saveProfiles({profiles:[{name:'客户',words:'张三\n北京某某科技有限公司'}],prefs:{mode:'code',ai:true}});
  const rec=await v.addRecord({kind:'doc',host:'wps',name:'合同-脱敏.docx',path:'C:\\a\\合同-脱敏.docx',mode:'code',counts:{WORD:2},total:2,pairs:[{original:'张三',masked:'【名称A】',type:'WORD',count:2},{original:'=1+1',masked:'x',type:'WORD',count:1}],data:{seq:{'st1#0':[['【名称A】','张三','WORD']]}}});
  const raw=await fsp.readFile(f,'utf8');assert.ok(!raw.includes('张三')&&!raw.includes('北京某某'));
  assert.equal((await v.findByPath('c:\\A\\合同-脱敏.docx'))[0].id,rec.id);
  const full=await v.getRecord(rec.id);assert.equal(full.data.seq['st1#0'][0][1],'张三');
  assert.equal(await v.currentWords(),'张三\n北京某某科技有限公司');
  const other=createMaskVault(f,{secrets:{available:true,async protect(){throw Error('x');},async unprotect(){throw Error('换了用户');}}});
  await assert.rejects(other.getRecord(rec.id),/原来的电脑/);
  const csv=await exportRecordCsv(full,d);const text=await fsp.readFile(csv,'utf8');assert.match(text,/^\ufeff原文,替换为,类型,次数/);assert.match(text,/张三,【名称A】,词表里的词,2/);assert.match(text,/'=1\+1/);
  assert.equal(await v.removeRecord(rec.id),true);await assert.rejects(v.getRecord(rec.id),/不在了/);
});
test('卡片和粘贴栏：选中的脱敏、对照表还原；还原前被改过的跳过',async t=>{
  const v=createMaskVault(null,{secrets:fakeSecrets}),cards=createCardStore(null),pb=createPasteBar(null);
  const c=await cards.add({type:'text',text:'张三 13812345678'},'wps'),p=await pb.add({type:'table',rows:[['姓名','电话'],['张三','13912345678']]},'et'),q=await pb.add({type:'text',text:'没有敏感内容'},'wps');
  const found=await scanModules({cards,pasteBar:pb,words:'张三'});same(found.map(x=>x.kind+x.count),['card2','paste2']);
  const r=await maskModules({cards,pasteBar:pb,vault:v,items:[{kind:'card',id:c.id},{kind:'paste',id:p.id},{kind:'paste',id:q.id}],mode:'star',words:'张三'});
  assert.equal(r.changed,2);assert.equal((await cards.list())[0].block.text,'张* 138****5678');same((await pb.list()).find(x=>x.id===p.id).block.rows[1],['张*','139****5678']);
  await pb.replaceBlock(p.id,{type:'text',text:'后来改过'});
  const back=await restoreModules({cards,pasteBar:pb,vault:v,id:r.record.id});assert.equal(back.restored,1);assert.match(back.skipped[0],/后来改过/);
  assert.equal((await cards.list())[0].block.text,'张三 13812345678');
});
test('AI 发送前：词表里的词也换成占位符，回答里换回来',async()=>{
  let sent='';const guard=createPrivacyGuard({getMode:async()=>'mask',getWords:async()=>'北京某某科技有限公司\n张三=员工A',plugins:[]});
  const r=await guard.beforeSend({messages:[{role:'user',content:'张三在北京某某科技有限公司，电话13812345678'}]});sent=r.messages.map(m=>m.content).join('\n');
  assert.ok(!sent.includes('张三')&&!sent.includes('北京某某'));assert.match(sent,/\[WORD_1\]/);assert.equal(r.report.masked,3);
  assert.equal(await guard.afterReceive(r.session,'已联系 [WORD_2]（[WORD_1]）'),'已联系 北京某某科技有限公司（张三）');
  const bad=createPrivacyGuard({getMode:async()=>'mask',getWords:async()=>{throw Error('坏了');},plugins:[]});await assert.rejects(bad.beforeSend({messages:[{role:'user',content:'x'}]}),/停止发送/);
});

// ---- WPS 文字：假的文档对象（正文 + 页眉两个文字流，段落以 \r 分隔）----
function fakeWord(stories){
  const st={};for(const [k,v] of Object.entries(stories))st[k]={text:v,ranges:[]};
  function range(story,s,e){const pos={s,e};story.ranges.push(pos);return {get Start(){return pos.s;},get End(){return pos.e;},get Text(){return story.text.slice(pos.s,pos.e);},set Text(v){const start=pos.s,end=pos.e,delta=v.length-(end-start);story.text=story.text.slice(0,start)+v+story.text.slice(end);for(const p of story.ranges){if(p.s>=end)p.s+=delta;if(p.e>=end)p.e+=delta;}pos.e=start+v.length;},get Duplicate(){return range(story,pos.s,pos.e);},SetRange(a,b){pos.s=a;pos.e=b;},
    get Paragraphs(){const list=[];let pos=0;story.text.split('\r').forEach((p,i,arr)=>{const end=pos+p.length+(i<arr.length-1?1:0);if(end>s&&pos<(e||1)||pos===s)list.push(range(story,pos,end));pos=end;});return {Count:list.length,Item:i=>({Range:list[i-1]})};}};}
  const d={Name:'合同.docx',FullName:'C:\\docs\\合同.docx',Saved:true,TrackRevisions:true,Revisions:{Count:0},savedAs:null,saves:0,vars:{},
    StoryRanges:{Item(t){if(!st[t])throw Error('none');return range(st[t],0,st[t].text.length);}},
    SaveAs2(p){this.savedAs=p;this.FullName=p;this.Name=p.split('\\').pop();},Save(){this.saves++;},UndoClear(){this.undoCleared=true;},
    Variables:{Item:n=>{if(!(n in d.vars))throw Error('x');return {Value:d.vars[n],Delete(){delete d.vars[n];}};},Add:(n,v)=>{d.vars[n]=v;}},
    BuiltInDocumentProperties:{props:{Author:'张三',Company:'某某公司','Last Author':'','Manager':''},Item(n){const P=this.props;return {get Value(){return P[n];},set Value(v){P[n]=v;}};}}};
  return {d,st,app:{ActiveDocument:d,Selection:{Range:range(st[1],0,0)}}};
}
test('WPS 文字：正文和页眉都改到，按位置改、核对原文，按顺序还原',async()=>{
  const ctx=browser(),E=ctx.MaskEngine,H=ctx.MaskHost,P=ctx.MaskPanel;
  const {d,st,app}=fakeWord({1:'甲方：张三\r电话13812345678，联系人张三\r',7:'张三 机密'});
  const ad=H.wpsWord(app),words=E.parseWords('张三').words;
  ad.copy('C:\\docs\\合同-脱敏.docx');assert.equal(d.savedAs,'C:\\docs\\合同-脱敏.docx');ad.prepare(true);assert.equal(d.TrackRevisions,false);
  const segs=ad.segments('all'),sess=E.session({mode:'block',words}),data={};
  same([...new Set(segs.map(s=>s.key))],['st1#0','st7#0']);
  for(const seg of segs){const hits=E.find(seg.text,{words});if(!hits.length)continue;const r=E.apply(seg.text,hits,sess);const ok=ad.write(seg,hits.map((h,i)=>({start:h.start,end:h.end,text:r.seq[i][0],expect:h.value})));r.seq.forEach((p,i)=>{if(ok[i])(data[seg.key]=data[seg.key]||[]).push(p);});}
  assert.equal(st[1].text,'甲方：■■\r电话■■■■■■■■■■■，联系人■■\r');assert.equal(st[7].text,'■■ 机密');
  ad.clean();assert.equal(d.BuiltInDocumentProperties.props.Author,'');ad.mark('rec1');assert.equal(ad.marked(),'rec1');
  // 还原：按记下的顺序找回去
  const segs2=ad.segments('all');for(const k of Object.keys(data)){const group=P.groupSegs(segs2,k),plan=P.restorePlan(group,data[k]);assert.equal(plan.missing,0);group.forEach((s,i)=>{if(plan.edits[i].length)ad.write(s,plan.edits[i]);});}
  assert.equal(st[1].text,'甲方：张三\r电话13812345678，联系人张三\r');assert.equal(st[7].text,'张三 机密');
  // 原文对不上（有域或被改过）就不改
  const seg=ad.segments('all')[0];same([...ad.write(seg,[{start:3,end:5,text:'XX',expect:'李四'}])],[false]);
  // 原文件有修订、又要直接改原文件：先停下来
  d.Revisions.Count=2;assert.throws(()=>ad.prepare(false),/修订/);
});
test('WPS 文字：四段变长与变短代号替换、部分选区和多段还原',()=>{
 const c=browser(),E=c.MaskEngine,original='张伟在华东商贸\r电话13812345678\r张伟的身份证11010519491231002X\r再联系张伟\r';
 const {app,st}=fakeWord({1:original}),ad=c.MaskHost.wpsWord(app),words=E.parseWords('张伟\n华东商贸').words;
 const segs=ad.segments('all'),session=E.session({mode:'code',words}),seq=[];
 const jobs=segs.map(seg=>{const hits=E.find(seg.text,{words}),r=E.apply(seg.text,hits,session);seq.push(...r.seq);return {seg,edits:hits.map((h,i)=>({start:h.start,end:h.end,expect:h.value,text:r.seq[i][0]}))};});
 assert.equal(seq.length,6);
 // Selection bounds are snapshots, evaluated against the original segment before any writes.
 const sample={...segs[1],lo:segs[1].start+2,hi:segs[1].start+13};
 assert.equal(ad.inside(sample,{start:2,end:13}),true);assert.equal(ad.inside(sample,{start:0,end:13}),false);
 jobs.forEach(j=>assert.ok(ad.write(j.seg,j.edits).every(Boolean)));
 assert.doesNotMatch(st[1].text,/张伟|华东商贸|13812345678|11010519491231002X/);
 const group=ad.segments('all'),plan=c.MaskPanel.restorePlan(group,seq);assert.equal(plan.missing,0);
 group.forEach((seg,i)=>assert.ok(ad.write(seg,plan.edits[i]).every(Boolean)));
 assert.equal(st[1].text,original);
});

test('WPS 表格：整格记原值，公式、数字格式原样还原；批注、页眉也处理',async()=>{
  const ctx=browser(),E=ctx.MaskEngine,H=ctx.MaskHost;
  const cells={'1,1':{v:'姓名',f:'姓名',nf:'General'},'2,1':{v:'张三',f:'张三',nf:'General'},'2,2':{v:13812345678,f:'13812345678',nf:'0'},'3,1':{v:'张三',f:'=A2',nf:'General'}};
  const cell=(r,c)=>{const k=r+','+c;cells[k]=cells[k]||{v:'',f:'',nf:'General'};const x=cells[k];return {get Value2(){return x.v;},set Value2(v){x.v=v;x.f=String(v);},get Formula(){return x.f;},set Formula(v){x.f=v;x.v='(算出)';},get NumberFormat(){return x.nf;},set NumberFormat(v){x.nf=v;}};};
  const note={t:'张三的电话',Text(v){if(v!==undefined)this.t=v;return this.t;},Parent:{Address:()=>'B2'}};
  const ps={LeftHeader:'',CenterHeader:'张三 专用',RightHeader:'',LeftFooter:'',CenterFooter:'',RightFooter:''};
  const used={Row:1,Column:1,Rows:{Count:3},Columns:{Count:2},get Value2(){return [1,2,3].map(r=>[1,2].map(c=>cell(r,c).Value2));},get Formula(){return [1,2,3].map(r=>[1,2].map(c=>cell(r,c).Formula));}};
  const ws={Name:'名单',ProtectContents:false,UsedRange:used,Cells:{Item:cell},Comments:{Count:1,Item:()=>note},PageSetup:ps};
  const wb={Name:'名单.xlsx',FullName:'C:\\a\\名单.xlsx',Saved:true,Worksheets:{Count:1,Item:()=>ws},CustomDocumentProperties:{props:{},Item(n){if(!(n in this.props))throw Error('x');const P=this.props;return {Value:P[n],Delete(){delete P[n];}};},Add(n,l,t,v){this.props[n]=v;}}};
  const ad=H.wpsSheet({ActiveWorkbook:wb,ActiveSheet:ws});
  const segs=ad.segments('all');same(segs.map(s=>s.key),['名单!R1C1','名单!R2C1','名单!R2C2','名单!R3C1','名单!note!B2','名单!ps!CenterHeader']);
  const words=E.parseWords('张三').words,sess=E.session({mode:'star',words}),saved={};
  segs.forEach(s=>{const hits=E.find(s.text,{words});if(!hits.length)return;const r=E.apply(s.text,hits,sess);saved[s.key]=ad.put(s,r.text);});
  assert.equal(cells['2,2'].v,'138****5678');assert.equal(cells['2,2'].nf,'@');assert.equal(cells['3,1'].v,'张*');assert.equal(note.t,'张*的电话');assert.equal(ps.CenterHeader,'张* 专用');
  same(ad.names(words),[]);ws.Name='张三的表';same(ad.names(words),['张三的表']);ws.Name='名单';
  const map={};ad.segments('all').forEach(s=>{map[s.key]=s;});Object.keys(saved).forEach(k=>assert.equal(ad.back(map[k],saved[k]),true));
  assert.equal(cells['2,2'].v,13812345678);assert.equal(cells['2,2'].nf,'0');assert.equal(cells['3,1'].f,'=A2');assert.equal(note.t,'张三的电话');assert.equal(ps.CenterHeader,'张三 专用');
  assert.equal(ad.mark('r9'),true);assert.equal(ad.marked(),'r9');
});
test('WPS 演示：文本框、表格、备注都改到，按字符位置核对',async()=>{
  const ctx=browser(),E=ctx.MaskEngine,H=ctx.MaskHost;
  const tr=t=>({t,get Text(){return this.t;},Characters(s,l){const self=this;return {get Text(){return self.t.substr(s-1,l);},set Text(v){self.t=self.t.slice(0,s-1)+v+self.t.slice(s-1+l);}};}});
  const box={Type:17,Id:3,HasTable:false,HasTextFrame:true,TextFrame:{HasText:true,TextRange:tr('联系人：张三\r电话：13812345678')}};
  const cellTr=tr('张三');const table={Type:19,Id:4,HasTable:true,Table:{Rows:{Count:1},Columns:{Count:1},Cell:()=>({Shape:{TextFrame:{TextRange:cellTr}}})}};
  const group={Type:6,Id:5,GroupItems:{Count:1,Item:()=>({Type:17,Id:6,HasTextFrame:true,TextFrame:{HasText:true,TextRange:tr('张三')}})}};
  const note={Id:2,HasTextFrame:true,PlaceholderFormat:{Type:2},TextFrame:{HasText:true,TextRange:tr('讲到张三时停一下')}};
  const slide={SlideID:256,Shapes:{Count:3,Item:i=>[box,table,group][i-1]},NotesPage:{Shapes:{Count:1,Item:()=>note}}};
  const pres={Name:'汇报.pptx',FullName:'C:\\a\\汇报.pptx',Saved:true,Slides:{Count:1,Item:()=>slide},Tags:{v:{},Add(n,x){this.v[n]=x;},Delete(n){delete this.v[n];},Item(n){return this.v[n]||'';}},CustomDocumentProperties:{Item(){throw Error('x');},Add(){throw Error('x');}}};
  const ad=H.wpsSlides({ActivePresentation:pres});const segs=ad.segments('all');
  same(segs.map(s=>s.key),['s256#3','s256#4/r1c1','s256#6','s256n#2']);
  const words=E.parseWords('张三').words,sess=E.session({mode:'code',words});
  segs.forEach(s=>{const hits=E.find(s.text,{words});const r=E.apply(s.text,hits,sess);ad.write(s,hits.map((h,i)=>({start:h.start,end:h.end,text:r.seq[i][0],expect:h.value})));});
  assert.equal(box.TextFrame.TextRange.t,'联系人：【名称A】\r电话：【电话1】');assert.equal(cellTr.t,'【名称A】');assert.equal(note.TextFrame.TextRange.t,'讲到【名称A】时停一下');
  assert.equal(ad.mark('p1'),true);assert.equal(ad.marked(),'p1');
});
test('服务接口：词表、副本文件名、对照表都要本机口令；不认识的动作 404',async t=>{
  const d=await tmp(t);await fsp.writeFile(path.join(d,'合同.docx'),'x');await fsp.writeFile(path.join(d,'合同-脱敏.docx'),'x');
  const server=createCoachServer({cardsFile:path.join(d,'c.json'),pasteBarFile:path.join(d,'p.json'),settingsFile:path.join(d,'s.json'),backupDir:d,aiFile:path.join(d,'ai.json'),maskFile:path.join(d,'masking.json'),maskSecrets:fakeSecrets});
  await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await (await fetch(base+'/api/session')).json()).token;
  const post=(b,k=token)=>fetch(base+'/api/tools/mask',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':k},body:JSON.stringify(b)});
  assert.equal((await post({action:'state'},'bad')).status,403);
  const st=await (await post({action:'state'})).json();assert.equal(st.protection,'dpapi');same(st.records,[]);
  const tg=await (await post({action:'target',source:path.join(d,'合同.docx'),name:'合同.docx'})).json();assert.equal(tg.path,path.join(d,'合同-脱敏-2.docx'));
  const back=await (await post({action:'target',source:path.join(d,'合同-脱敏.docx'),name:'合同-脱敏.docx',suffix:'restore'})).json();assert.equal(back.path,path.join(d,'合同-还原.docx'));
  const ja=await (await post({action:'target',source:path.join(d,'合同.docx'),name:'合同.docx',lang:'ja'})).json();assert.equal(path.basename(ja.path),'合同-マスキング.docx');
  const saved=await (await post({action:'save',record:{kind:'doc',host:'wps',name:'合同-脱敏-2.docx',path:tg.path,mode:'star',counts:{PHONE:1},total:1,pairs:[{original:'13812345678',masked:'138****5678',type:'PHONE',count:1}],data:{seq:{}}}})).json();
  assert.equal((await (await post({action:'find',path:tg.path})).json()).records[0].id,saved.record.id);
  assert.ok(!(await fsp.readFile(path.join(d,'masking.json'),'utf8')).includes('13812345678'));
  assert.equal((await post({action:'nope'})).status,404);
  // 卡片自动脱敏：开着时存进去就脱敏
  await post({action:'profiles',profiles:[{name:'a',words:'张三'}],prefs:{autoModules:true,mode:'block'}});
  const card=await (await fetch(base+'/api/cards',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify({block:{type:'text',text:'张三 13812345678'},sourceHost:'wps'})})).json();
  assert.equal(card.card.block.text,'■■ ■■■■■■■■■■■');
});
test('数据脱敏：功能区、侧栏、Office 版、日文词条都接上了',()=>{
  const root=f=>fs.readFileSync(new URL('../'+f,import.meta.url),'utf8');
  assert.match(read('ribbon.xml'),/id="dataMask"/);assert.match(read('ribbon-text.js'),/"dataMask":\{"label":"数据脱敏"/);assert.match(read('tool-icons.js'),/dataMask:'shield'/);assert.ok(fs.existsSync(new URL('../public/icons/shield.svg',import.meta.url)));
  assert.match(read('main.js'),/dataMask:'jumpMask'/);assert.match(read('quick-panel.js'),/id:'jumpMask'/);assert.match(read('tools.html'),/mask-panel\.js/);assert.match(read('office/taskpane.html'),/\.\.\/mask-panel\.js/);
  assert.match(root('lib/office-server.mjs'),/'\/mask-engine\.js','\/mask-host\.js','\/mask-panel\.js'/);assert.match(read('office/office-batch2.js'),/TOOLS\.word\.dataMask=maskJump/);assert.match(read('i18n-ja-extra.js'),/"数据脱敏":"データのマスキング"/);
});
test('功能区「常用工具」：文字、表格、演示各放自己的常用功能，公文格式按界面语言换成ビジネス文書',()=>{
  const xml=read('ribbon.xml'),g=xml.slice(xml.indexOf('<group id="toolsGroup"'),xml.indexOf('</group>',xml.indexOf('<group id="toolsGroup"')));
  const ids=[...g.matchAll(/<button[^>]*id="([^"]+)"/g)].map(m=>m[1]);
  same(ids.filter(i=>i.startsWith('cw_')),['cw_bodyFormat','cw_officialFormatCn','cw_businessFormatJa','cw_deleteBlankPages','cw_proofCheck','cw_jaStyle','cw_exportPdf']);
  same(ids.filter(i=>i.startsWith('ce_')),['ce_spotlight','ce_sheetCheck','ce_formulaOpen','ce_removeDuplicates','ce_textToNumber','ce_exportPdf']);
  same(ids.filter(i=>i.startsWith('cp_')),['cp_studioOpen','cp_unifyDeckFont','cp_applyDesign','cp_slideNumbers','cp_slidesToImages','cp_slimDeck']);
  // 共用的组放在各程序自己的组后面，并收进菜单
  const order=[...xml.matchAll(/<group id="([^"]+)"/g)].map(m=>m[1]);assert.ok(order.indexOf('fileGroup')>order.indexOf('slideOutGroup'));assert.ok(order.indexOf('favGroup')===1);
  assert.match(xml,/<menu id="transferMenu"[^>]*size="large">/);assert.match(xml,/<menu id="officeMenu"[^>]*size="large">/);
  // main.js：按程序和语言显示，文字和动作用原来的按钮
  const calls=[];const ctx={location:{origin:'http://x'},setInterval(){}};ctx.window=ctx;let host='wps',lng='zh',hidden=[];
  ctx.OfficeTools={host:()=>host,isHidden:id=>hidden.includes(id),list:()=>[],app:()=>({}),runFromUi:id=>{calls.push(id);return Promise.resolve('ok');},version:'x'};
  ctx.OfficeI18n={t:x=>x,lang:()=>lng};ctx.OfficeRibbonText={bodyFormat:{label:'正文排版',tip:'t'}};ctx.ToolIcons={forAction:id=>'icons/'+id+'.svg'};ctx.PaneBridge={view:v=>v};ctx.fetch=undefined;
  vm.runInNewContext(read('main.js'),ctx);
  const V=id=>ctx.GetVisible({Id:id});
  assert.equal(V('cw_bodyFormat'),true);assert.equal(V('ce_spotlight'),false);assert.equal(V('cw_officialFormatCn'),true);assert.equal(V('cw_businessFormatJa'),false);
  lng='ja';assert.equal(V('cw_officialFormatCn'),false);assert.equal(V('cw_businessFormatJa'),true);
  hidden=['bodyFormat'];assert.equal(V('cw_bodyFormat'),false);hidden=[];
  host='wpp';assert.equal(V('cp_slimDeck'),true);assert.equal(V('cw_bodyFormat'),false);assert.equal(V('studioOpen'),false);host='et';assert.equal(V('studioOpen'),true);
  assert.equal(ctx.GetLabel({Id:'cw_bodyFormat'}),'正文排版');assert.equal(ctx.GetImage({Id:'ce_spotlight'}),'icons/spotlight.svg');
});
test('GitHub 自动测试：三个系统跑测试，Mac 上装一次、检查、卸载；日文安装包能打',()=>{
  const root=f=>fs.readFileSync(new URL('../'+f,import.meta.url),'utf8');
  // 工作流文件的“原件”放在 scripts/ci/workflows（远程工具不能写 .github）；上 GitHub 时复制到 .github/workflows，两边要一样。
  for(const f of ['test.yml','release-mac.yml']){const gh=new URL('../.github/workflows/'+f,import.meta.url);if(fs.existsSync(gh))assert.equal(fs.readFileSync(gh,'utf8'),root('scripts/ci/workflows/'+f),'.github/workflows/'+f+' 和 scripts/ci/workflows 里的不一样');}
  const wf=root('scripts/ci/workflows/test.yml');assert.match(wf,/windows-latest, macos-latest, ubuntu-latest/);assert.match(wf,/scripts\/ci\/secret-check\.mjs/);assert.match(wf,/bash scripts\/ci\/mac-install\.sh/);assert.match(wf,/mac-check\.mjs uninstall/);
  assert.match(root('scripts/ci/workflows/release-mac.yml'),/build-release\.mjs --base \. --out release --only mac/);
  for(const f of ['scripts/ci/secret-check.mjs','scripts/ci/mac-check.mjs','scripts/ci/mac-install.sh'])assert.ok(fs.existsSync(new URL('../'+f,import.meta.url)),f);
  // 日文包把默认语言改成 ja：settings.mjs 里的写法要能对上
  assert.ok(root('lib/settings.mjs').includes("ui:{language:'auto'"));assert.match(root('scripts/package/build-release.mjs'),/s\.replace\("ui:\{language:'auto'","ui:\{language:'ja'"\)/);
  // Mac 安装：Node 要 22 以上
  assert.match(root('scripts/package-mac/install.command'),/"\$major" -ge 22/);
});
