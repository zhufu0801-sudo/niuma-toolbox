import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
const code=toolsCode();
function sheet(grid){ // grid: rows of values; undefined = blank
  const cells=grid.map(row=>row.map(v=>({Value2:v,HasFormula:false})));
  const range={Rows:{Count:grid.length},Columns:{Count:grid[0].length},MergeCells:false,WrapText:false,Cells:{Item:(r,c)=>cells[r-1][c-1]}};
  const ws={Name:'S1'},wb={Name:'a.xlsx',FullName:'C:/a.xlsx',ActiveSheet:ws};
  range.Address=()=>'$A$1';range.Worksheet=ws;
  return {cells,range,app:{ActiveWorkbook:wb,ActiveSheet:ws,Selection:range}};
}
function run(app,id,answer=true){const window={Application:app,confirm:()=>answer};vm.runInNewContext(code,{window,fetch:async()=>{throw Error('no network');},location:{pathname:'/addon/et/'},setTimeout,clearTimeout});return window.OfficeTools.run(id,{host:'et'});}
test('跨列居中：只允许左边一格有内容，设置后不合并',async()=>{
  const ok=sheet([['标题',undefined,undefined]]);assert.match(await run(ok.app,'centerAcross'),/跨列居中/);assert.equal(ok.range.HorizontalAlignment,7);
  const bad=sheet([['标题','备注',undefined]]);await assert.rejects(run(bad.app,'centerAcross'),/多个非空/);assert.equal(bad.range.HorizontalAlignment,undefined);
  const m=sheet([['标题',undefined]]);m.range.MergeCells=true;await assert.rejects(run(m.app,'centerAcross'),/合并/);
});
test('缩小字体填充：关闭换行后开启，值不变',async()=>{
  const s=sheet([['很长的文字']]);s.range.WrapText=true;assert.match(await run(s.app,'shrinkToFit'),/换行已关闭/);assert.equal(s.range.ShrinkToFit,true);assert.equal(s.range.WrapText,false);assert.equal(s.cells[0][0].Value2,'很长的文字');
});
test('空白填上一格：只填真正空白，顶端空白跳过，公式与空格不动',async()=>{
  const s=sheet([[undefined,'甲'],['销售',undefined],[undefined,' '],['财务',undefined]]);s.cells[2][0].HasFormula=true;s.cells[2][0].Value2='';
  assert.match(await run(s.app,'fillBlanks'),/已填充 2 个/);
  assert.deepEqual(s.cells.map(r=>r.map(c=>c.Value2)),[[undefined,'甲'],['销售','甲'],['',' '],['财务',' ']]);
  const c=sheet([['A'],[undefined]]);assert.match(await run(c.app,'fillBlanks',false),/已取消/);assert.equal(c.cells[1][0].Value2,undefined);
});
