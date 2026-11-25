import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
const code=fs.readFileSync(new URL('../public/host.js',import.meta.url),'utf8');
function read(host,application){const window={Application:application};vm.runInNewContext(code,{window});return window.CoachBridge.read(host);}
test('caps Word selection and does not read full document content',()=>{
  const result=read('wps',{ActiveDocument:{Name:'test.docx',get Content(){throw new Error('must not read document body');}},Selection:{Text:'中文'.repeat(5000)}});
  assert.equal(result.selection.length,4000);assert.equal(result.document,'test.docx');assert.equal(result.connected,true);
});
test('reads at most 40 spreadsheet cells even for a whole-sheet selection',()=>{
  let reads=0;const result=read('et',{ActiveWorkbook:{Name:'test.xlsx'},ActiveSheet:{Name:'Sheet1'},Selection:{Address:'$A:$XFD',Rows:{Count:1048576},Columns:{Count:16384},Cells:{Item(r,c){reads++;return {Text:`${r},${c}`};}}}});
  assert.equal(reads,40);assert.equal(result.selection.split('\n').length,8);assert.match(result.detail,/8 行 × 5 列/);
});
test('PPT uses current slide text when no text is selected',()=>{
  const result=read('wpp',{ActivePresentation:{Name:'test.pptx'},ActiveWindow:{Selection:{},View:{Slide:{SlideIndex:3,Shapes:{Count:2,Item(i){return {TextFrame:{TextRange:{Text:'段落'+i}}};}}}}}});
  assert.match(result.location,/第 3 页/);assert.equal(result.selection,'段落1\n段落2');
});
test('browser preview never claims to have read a document',()=>{assert.equal(read('wps',undefined).connected,false);});
