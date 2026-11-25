import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {createCoachServer} from '../server.mjs';
import {toolsCode} from './helpers/tools-code.mjs';
const code=toolsCode();
function bridge(application,fetchImpl=async()=>{throw Error('unexpected network');}){const window={Application:application,confirm:()=>true};vm.runInNewContext(code,{window,fetch:fetchImpl,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date});return window.OfficeTools;}
const doc=()=>({Name:'test.docx',FullName:'C:/test.docx'});
test('missing WPS, empty selection, read-only documents and oversized spreadsheet selection are rejected',async()=>{
  await assert.rejects(bridge().run('heading1'),/请在 WPS/);
  await assert.rejects(bridge({ActiveDocument:doc(),Selection:{Text:' '}}).run('heading1'),/选中/);
  await assert.rejects(bridge({ActiveDocument:{...doc(),ReadOnly:true}}).run('heading1'),/只读/);
  await assert.rejects(bridge({ActiveWorkbook:{Name:'a.xlsx'},Selection:{Rows:{Count:1001},Columns:{Count:2}}}).run('fitColumns',{host:'et'}),/1000/);
});
test('heading changes the selected range and closes its undo group',async()=>{
  const calls=[],range={};const a={ActiveDocument:doc(),Selection:{Text:'标题',Range:range},UndoRecord:{StartCustomRecord(){calls.push('start');},EndCustomRecord(){calls.push('end');}}};
  await bridge(a).run('heading1');assert.equal(range.Style,-2);assert.deepEqual(calls,['start','end']);
});
test('Word table inserts after selection, never overwrites it',async()=>{
  let collapsed,where;const selectionRange={Duplicate:{Collapse(mode){collapsed=mode;}}};const values=[];
  const d={...doc(),Tables:{Add(range,rows,cols){where=range;assert.equal(rows,2);assert.equal(cols,2);return {Cell(r,c){return {Range:{set Text(v){values.push([r,c,v]);}}};},Borders:{}};}}};
  await bridge({ActiveDocument:d,Selection:{Range:selectionRange}}).run('insertTable',{text:'姓名\t任务\n小林\t排版'});
  assert.equal(collapsed,0);assert.equal(where,selectionRange.Duplicate);assert.equal(values[3][2],'排版');
});
function slideApp(fail=false){const created=[],deleted=[];const d={Name:'sample.pptx',PageSetup:{SlideWidth:960,SlideHeight:540},Slides:{Count:0,Add(index){const slide={Shapes:{AddTextbox(){if(fail)throw Error('API failure');return {TextFrame:{TextRange:{Font:{}}}};}},Delete(){deleted.push(index);}};created.push(slide);this.Count++;return slide;}}};return {a:{ActivePresentation:d},created,deleted};}
test('PPT creates actual slide objects and rolls new slides back on API failure',async()=>{
  const good=slideApp();assert.match(await bridge(good.a).run('makeSlides',{host:'wpp',text:'项目概览\n目标\n\n计划\n任务'}),/2 页/);assert.equal(good.created.length,2);
  const bad=slideApp(true);await assert.rejects(bridge(bad.a).run('makeSlides',{host:'wpp',text:'标题\n正文'}),/已回退/);assert.deepEqual(bad.deleted,[1]);
});
function fakeFetch(block,hook){return async url=>{if(url==='/api/session')return Response.json({token:'abc'});if(hook)hook();return Response.json({block,sourceHost:'wps'});};}
test('asynchronous transfer stops if the active document changed',async()=>{
  const a={ActiveDocument:doc(),Selection:{Range:{Start:1,End:4}}};const tool=bridge(a,fakeFetch({type:'text',text:'内容'},()=>{a.ActiveDocument={Name:'other.docx',FullName:'C:/other.docx'};}));
  await assert.rejects(tool.run('insertTransfer'),/已切换/);
});
test('Excel transfer refuses to overwrite existing cells',async()=>{
  let writes=0;const destination={MergeCells:false,set Value2(v){writes++;}},range={Rows:{Count:1},Columns:{Count:1},Cells:{Item(){return {Resize(){return destination;}};}}};
  const a={ActiveWorkbook:{Name:'sample.xlsx'},ActiveSheet:{},Selection:range,WorksheetFunction:{CountA(){return 1;}}};
  await assert.rejects(bridge(a,fakeFetch({type:'text',text:'内容'})).run('insertTransfer',{host:'et'}),/避免覆盖/);assert.equal(writes,0);
});
test('Excel treats formula-looking transferred content as literal text',async()=>{
  let written;const destination={MergeCells:false,set Value2(value){written=value;}},range={Rows:{Count:1},Columns:{Count:1},Cells:{Item(){return {Resize(){return destination;}};}}};
  const a={ActiveWorkbook:{Name:'sample.xlsx'},ActiveSheet:{},Selection:range,WorksheetFunction:{CountA(){return 0;}}};
  await bridge(a,fakeFetch({type:'table',rows:[['=1+1','正常文字']]})).run('insertTransfer',{host:'et'});
  assert.equal(written[0][0],"'=1+1");assert.equal(written[0][1],'正常文字');
});
test('canceling PDF folder selection never exports a document',async()=>{
  let exported=false;const a={ActiveDocument:{...doc(),ExportAsFixedFormat(){exported=true;}},FileDialog(){return {Show(){return 0;}};}};
  assert.match(await bridge(a).run('exportPdf'),/已取消/);assert.equal(exported,false);
});
test('switching sheets during transfer confirmation prevents insertion',async()=>{
  let written=false;const a={ActiveWorkbook:{Name:'test.xlsx'},ActiveSheet:{Name:'Sheet1'},Selection:{Rows:{Count:1},Columns:{Count:1},Address:'A1'}};
  const window={Application:a,confirm(){a.ActiveSheet={Name:'Sheet2'};return true;}};
  vm.runInNewContext(code,{window,fetch:fakeFetch({type:'text',text:'内容'}),location:{pathname:'/addon/et/',search:''},URLSearchParams,Date});
  await assert.rejects(window.OfficeTools.run('insertTransfer',{host:'et'}),/工作表或选区已切换/);
});
test('PPT non-text selection gives actionable instruction and empty outline does not claim rollback',async()=>{
  const a={ActivePresentation:{Name:'test.pptx'},ActiveWindow:{Selection:{Type:2,get TextRange(){throw Error('must not inspect non-text');}}}};
  await assert.rejects(bridge(a).run('capture',{host:'wpp'}),/选中幻灯片里的文字/);
  const blank=slideApp();await assert.rejects(bridge(blank.a).run('makeSlides',{host:'wpp',text:''}),e=>/没有可生成/.test(e.message)&&!/回退/.test(e.message));
});
test('local transfer API enforces session, empty state, size and explicit clear',async t=>{
  const server=createCoachServer();await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await(await fetch(base+'/api/session')).json()).token;
  const invoke=(method,body,auth=token)=>fetch(base+'/api/local-transfer',{method,headers:{'Content-Type':'application/json','X-Coach-Token':auth},body:body?JSON.stringify(body):undefined});
  assert.equal((await invoke('GET',null,'wrong')).status,403);assert.equal((await invoke('GET')).status,404);
  const payload={sourceHost:'wps',block:{type:'text',text:'选中内容'}};assert.equal((await invoke('POST',payload)).status,200);assert.equal((await(await invoke('GET')).json()).block.text,'选中内容');
  assert.equal((await invoke('POST',{sourceHost:'et',block:{type:'table',rows:Array(21).fill(['a'])}})).status,422);
  assert.equal((await invoke('DELETE')).status,200);assert.equal((await invoke('GET')).status,404);
  const health=await(await fetch(base+'/api/health')).json();assert.equal(health.aiEnabled,false);assert.equal((await fetch(base+'/api/ask',{method:'POST'})).status,410);
});
