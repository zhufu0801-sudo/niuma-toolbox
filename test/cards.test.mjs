// 内容卡片：服务端存取、插入三种文档时的适配、拖入内容的识别。
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import {createCardStore,CARD_LIMIT} from '../lib/cards.mjs';
import {createCoachServer} from '../server.mjs';
import {toolsCode} from './helpers/tools-code.mjs';

test('卡片保存到文件，重启后还在；超过上限丢掉最旧的；内容会校验',async()=>{
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'cards-'));const file=path.join(dir,'cards.json');
  const a=createCardStore(file);
  const first=await a.add({type:'text',text:'第一张卡片\n第二行'},'wps');
  assert.equal(first.title,'第一张卡片');
  await a.add({type:'table',rows:[['姓名','部门'],['小林']]},'et');
  const b=createCardStore(file),list=await b.list();
  assert.equal(list.length,2);assert.equal(list[0].block.type,'table');assert.deepEqual(list[0].block.rows[1],['小林','']);
  for(let i=0;i<CARD_LIMIT+5;i++)await b.add({type:'text',text:'卡'+i},'drop');
  assert.equal((await b.list()).length,CARD_LIMIT);
  await assert.rejects(b.add({type:'text',text:'  '},'wps'),/空的/);
  await assert.rejects(b.add({type:'table',rows:Array(21).fill(['a'])},'wps'),/20 行/);
  await assert.rejects(b.add({type:'text',text:'x'},'evil'),/来源/);
  assert.ok(await b.remove((await b.list())[0].id));assert.equal(await b.remove('none'),false);
  await b.clear();assert.equal((await createCardStore(file).list()).length,0);
});

test('卡片接口需要本机令牌；放入中转区的内容也会出现在卡片里',async t=>{
  const server=createCoachServer();await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>server.close());
  const base='http://127.0.0.1:'+server.address().port;const {token}=await (await fetch(base+'/api/session')).json();
  const call=(method,body,auth=token,q='')=>fetch(base+'/api/cards'+q,{method,headers:{'Content-Type':'application/json','X-Coach-Token':auth},body:body?JSON.stringify(body):undefined});
  assert.equal((await call('GET',null,'wrong')).status,403);
  assert.equal((await call('POST',{block:{type:'text',text:'你好'},sourceHost:'wpp'})).status,200);
  await fetch(base+'/api/local-transfer',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify({block:{type:'text',text:'中转来的'},sourceHost:'wps'})});
  const {cards}=await (await call('GET')).json();
  assert.deepEqual(cards.map(c=>c.block.text),['中转来的','你好']);
  assert.equal((await call('DELETE',null,token,'?id='+cards[0].id)).status,200);
  assert.equal((await call('DELETE',null,token,'?id=none')).status,404);
  assert.equal((await call('POST',{block:{type:'text',text:'x'.repeat(10001)},sourceHost:'wps'})).status,422);
});

function bridge(a,cards=[],host='wps'){const window={Application:a,confirm:()=>true};
  const fetchImpl=async url=>url==='/api/session'?Response.json({token:'t'}):Response.json({cards});
  vm.runInNewContext(toolsCode(),{window,fetch:fetchImpl,location:{pathname:'/addon/'+host+'/',search:''},URLSearchParams,Date,setTimeout,clearTimeout});return window.OfficeTools;}
const doc=()=>({Name:'a.docx',FullName:'C:/a.docx'});

test('插入卡片到文字：表格插在选区末尾并适应页面宽度',async()=>{
  let autofit,where;const sel={Range:{Start:3,End:3,Duplicate:{Collapse(){} }}};
  const d={...doc(),Tables:{Add(r,rows,cols){where=r;return {Cell:()=>({Range:{}}),Borders:{},AutoFitBehavior(m){autofit=m;}};}}};
  const msg=await bridge({ActiveDocument:d,Selection:sel}).run('insertCard',{block:{type:'table',rows:[['a','b'],['c','d']]}});
  assert.equal(autofit,2);assert.ok(where);assert.match(msg,/卡片内容/);
});

test('插入卡片到表格：目标区域有内容时拒绝，避免覆盖',async()=>{
  let writes=0;const destination={MergeCells:false,set Value2(v){writes++;}},range={Rows:{Count:1},Columns:{Count:1},Cells:{Item(){return {Resize(){return destination;}};}}};
  const a={ActiveWorkbook:{Name:'s.xlsx'},ActiveSheet:{},Selection:range,WorksheetFunction:{CountA:()=>1}};
  await assert.rejects(bridge(a,[],'et').run('insertCard',{host:'et',block:{type:'text',text:'x'}}),/避免覆盖/);assert.equal(writes,0);
});

test('插入最新卡片：没有卡片时提示先存；卡片内容异常会被拦下',async()=>{
  const sel={Range:{Start:1,End:1}};
  await assert.rejects(bridge({ActiveDocument:doc(),Selection:sel}).run('insertLatest'),/还没有卡片/);
  await assert.rejects(bridge({ActiveDocument:doc(),Selection:sel}).run('insertCard',{block:{type:'script',text:'x'}}),/无法识别/);
  await assert.rejects(bridge({ActiveDocument:doc(),Selection:sel}).run('insertCard',{block:{type:'table',rows:[Array(11).fill('a')]}}),/10 列/);
});

test('拖入内容识别：HTML 表格、Tab 分隔表格、普通文字，超出范围时标记截断',()=>{
  const ui=fs.readFile(new URL('../public/tools-ui.js',import.meta.url),'utf8');
  return ui.then(code=>{
    const el=()=>({hidden:false,append(){},classList:{add(){},remove(){},toggle(){}},textContent:'',style:{},setAttribute(){},dataset:{},remove(){},querySelectorAll:()=>[]});
    const els={};const document={getElementById:id=>els[id]||(els[id]=el()),querySelectorAll:()=>[],createElement:()=>el()};
    const fakeTable=rows=>({rows:rows.map(r=>({cells:r.map(c=>({innerText:c}))}))});
    class DOMParser{parseFromString(html){return {querySelector:()=>html.includes('<table')?fakeTable([['姓名','部门'],['小林','设计']]):null};}}
    const window={};const OfficeTools={host:()=>'wps',list:()=>[],hidden:()=>[],isHidden:()=>false,app:()=>({}),cards:{list:()=>Promise.resolve([])}};
    vm.runInNewContext(code,{window,document,OfficeTools,ToolIcons:{path:()=>''},PaneBridge:{listen(){}},StudioLaunch:{},location:{search:'',origin:'o'},URLSearchParams,history:{replaceState(){}},setInterval(){},DOMParser,Date});
    const read=OfficeTools.cardsUi.readDrop,dt=(html,text)=>({getData:k=>k==='text/html'?html:k==='text/plain'?text:''});
    assert.equal(JSON.stringify(read(dt('<table>…</table>','x')).block),JSON.stringify({type:'table',rows:[['姓名','部门'],['小林','设计']]}));
    assert.equal(read(dt('','甲\t乙\n丙\t丁\n')).block.rows[1][1],'丁');
    assert.equal(read(dt('','只是一段文字\n第二行')).block.type,'text');
    assert.equal(read(dt('','   ')),null);
    assert.equal(read(dt('','a\n'.repeat(1)+'b'.repeat(10001))).cut,true);
    const big=Array.from({length:25},(_,i)=>i+'\tx').join('\n');assert.equal(read(dt('',big)).cut,true);assert.equal(read(dt('',big)).block.rows.length,20);
    assert.equal(OfficeTools.cardsUi.htmlOf({type:'text',text:'<b>'}),'<p>&lt;b&gt;</p>');
  });
});
