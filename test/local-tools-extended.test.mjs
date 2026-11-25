import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
const code=toolsCode();
function bridge(a,confirm=()=>true){const window={Application:a,confirm};vm.runInNewContext(code,{window,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date});return window.OfficeTools;}
function sheetApp(cells){const range={MergeCells:false,Address:'A1:C1',Rows:{Count:1},Columns:{Count:cells.length},Cells:{Item(r,c){return cells[c-1];}}};return {ActiveWorkbook:{Name:'test.xlsx'},ActiveSheet:{Name:'Sheet1'},Selection:range};}
test('spreadsheet text cleanup skips formulas and numeric values, preserving text identities',async()=>{
  const a={Value2:'  00123  ',HasFormula:false},b={Value2:123,HasFormula:false},c={Value2:'=SUM(A1)',HasFormula:true};
  const result=await bridge(sheetApp([a,b,c])).run('trimCells',{host:'et'});
  assert.match(result,/已修改 1/);assert.equal(a.Value2,"'00123");assert.equal(b.Value2,123);assert.equal(c.Value2,'=SUM(A1)');
});
test('literal replacement does not interpret regex or formula syntax',async()=>{
  const cell={Value2:'[a].[a]',HasFormula:false};await bridge(sheetApp([cell])).run('replaceText',{host:'et',findText:'[a]',replaceText:'=1'});assert.equal(cell.Value2,"'=1.=1");
});
test('spreadsheet text cleanup restores original cells if a later write fails',async()=>{
  const first={Value2:' a ',HasFormula:false};let second=' b ',fail=true;
  const other={HasFormula:false,get Value2(){return second;},set Value2(v){if(fail){fail=false;throw Error('read-only cell');}second=v;}};
  await assert.rejects(bridge(sheetApp([first,other])).run('trimCells',{host:'et'}),/已恢复/);assert.equal(first.Value2,"' a ");assert.equal(second,"' b ");
});
test('Word literal replacement uses backward ranges and preserves outside selection',async()=>{
  let content='前甲甲后';const range={Start:1,End:3,StoryType:1,Text:'甲甲',Tables:{Count:0},Fields:{Count:0}};
  const doc={Name:'test.docx',Range(start,end){return {set Text(value){content=content.slice(0,start)+value+content.slice(end);}};}};
  await bridge({ActiveDocument:doc,Selection:{Text:'甲甲',Range:range}}).run('replaceText',{findText:'甲',replaceText:'乙乙'});assert.equal(content,'前乙乙乙乙后');
});
test('Word replace rejects protected structures and cancellation makes no edits',async()=>{
  let writes=0;const range={Start:0,End:2,StoryType:1,Text:'甲乙',Tables:{Count:1},Fields:{Count:0}},doc={Name:'test.docx',Range(){writes++;}};
  const a={ActiveDocument:doc,Selection:{Text:'甲乙',Range:range}};
  await assert.rejects(bridge(a).run('replaceText',{findText:'甲',replaceText:'丙'}),/普通文字/);
  range.Tables.Count=0;assert.match(await bridge(a,()=>false).run('replaceText',{findText:'甲',replaceText:'丙'}),/取消/);assert.equal(writes,0);
});
test('blank paragraph removal skips partial and final document paragraphs',async()=>{
  const deleted=[];const paragraphs=[{Start:0,End:4,Text:'文字\r'},{Start:4,End:5,Text:'\r'},{Start:5,End:6,Text:'\r'},{Start:6,End:7,Text:'\r'}];paragraphs.forEach((p,i)=>p.Delete=()=>deleted.push(i));
  const range={Start:1,End:7,StoryType:1,Tables:{Count:0},Paragraphs:{Count:4,Item(i){return {Range:paragraphs[i-1]};}}};
  await bridge({ActiveDocument:{Name:'test.docx',Content:{End:7}},Selection:{Text:'字\r\r\r',Range:range}}).run('removeBlankParas');assert.deepEqual(deleted,[2,1]);
});
test('PPT duplicate targets only the current slide',async()=>{
  let index,duplicated=0;const a={ActivePresentation:{Name:'test.pptx',Slides:{Range(i){index=i;return {Duplicate(){duplicated++;}};}}},ActiveWindow:{View:{Slide:{SlideIndex:4}}}};
  await bridge(a).run('duplicateSlide',{host:'wpp'});assert.equal(index,4);assert.equal(duplicated,1);
});
test('table paste retains empty edge cells and text whitespace',()=>{
  const rows=bridge().tableInput({text:'\t任务\n小林\t排版\t'});assert.equal(rows[0][0],'');assert.equal(rows[0][1],'任务');assert.equal(rows[1].length,3);assert.equal(rows[1][2],'');
});
