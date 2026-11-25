import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
import {toolsCode} from './helpers/tools-code.mjs';
const w={};vm.runInNewContext(fs.readFileSync(new URL('../public/locale-engine.js',import.meta.url),'utf8'),{window:w});const E=w.LocaleEngine;
test('人民币大写：常见金额、零的读法、角分、负数、非法输入',()=>{
  const cases={'1234.56':'壹仟贰佰叁拾肆元伍角陆分',1005:'壹仟零伍元整',100000:'壹拾万元整',10010:'壹万零壹拾元整',100000001:'壹亿零壹元整',0.05:'伍分',10.05:'壹拾元零伍分',0:'零元整','-8.5':'负捌元伍角','¥1,200.00':'壹仟贰佰元整',120000000:'壹亿贰仟万元整',1010101:'壹佰零壹万零壹佰零壹元整'};
  for(const [k,v] of Object.entries(cases))assert.equal(E.rmbUpper(isNaN(Number(k))?k:Number(k)),v,k);
  assert.equal(E.rmbUpper('abc'),null);assert.equal(E.rmbUpper(1e13),null);
});
test('身份证：校验位、生日、性别、15 位旧号码',()=>{
  const base='11010519900307123';const c='10X98765432'[[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2].reduce((s,x,i)=>s+x*Number(base[i]),0)%11];
  assert.equal(JSON.stringify(E.idInfo(base+c)),JSON.stringify({birth:'1990-03-07',gender:'男'}));
  assert.match(E.idInfo(base+(c==='0'?'1':'0')).error,/校验位/);assert.match(E.idInfo('123').error,/18 位/);assert.equal(E.idInfo('110105900307122').gender,'女');
});
test('手机号、全角半角、邮编',()=>{
  assert.equal(E.phoneCheck('+86 138-1234-5678'),'有效：138 1234 5678');assert.match(E.phoneCheck('1381234567'),/位数不对/);assert.match(E.phoneCheck('23812345678'),/号段/);
  assert.equal(E.zenHan('ＡＢＣ１２３　ｶﾀｶﾅ ﾃﾞｰﾀ（株）'),'ABC123　カタカナ データ（株）');assert.equal(E.zenHan('普通の文'),null);assert.equal(E.zenHan('Ab9','full'),'Ａｂ９');assert.equal(E.zenHan('价格：１００，含税！'),'价格：100，含税！');
  assert.equal(E.postal('〒１２３４５６７'),'123-4567');assert.equal(E.postal('123-4567'),null);assert.equal(E.postal('100 080'),'100080');assert.equal(E.postal('12'),null);
});
test('西历和历互转：年号边界、元年、日期序列号',()=>{
  assert.equal(E.toWareki('2026/11/12'),'令和8年11月12日');assert.equal(E.toWareki('2019-05-01'),'令和元年5月1日');assert.equal(E.toWareki('2019/4/30'),'平成31年4月30日');
  assert.equal(E.toWareki(46338),'令和8年11月12日');assert.equal(E.toWareki('1800/1/1'),null);
  assert.equal(E.toSeireki('令和8年11月12日').text,'2026/11/12');assert.equal(E.toSeireki('R8.11.12').serial,46338);assert.equal(E.toSeireki('平成元年1月8日').text,'1989/1/8');
  assert.match(E.toSeireki('平成31年5月1日').error,/范围/);assert.match(E.toSeireki('令和1年2月30日').error,/不存在/);assert.equal(E.toSeireki('2026/1/1'),null);
});
function sheet(grid){const cells=grid.map(row=>row.map(v=>({Value2:v,HasFormula:false})));const right=grid.map(row=>row.map(()=>({Value2:undefined}))).map(r=>r.concat([{Value2:undefined}]));
  const range={Rows:{Count:grid.length},Columns:{Count:grid[0].length},Cells:{Item:(r,c)=>cells[r-1][c-1]},Offset:(r,c)=>({Cells:{Item:(i,j)=>right[i-1][j-1]}})};
  const ws={Name:'S1'},wb={Name:'a.xlsx',FullName:'C:/a.xlsx',ActiveSheet:ws};range.Address=()=>'A1';range.Worksheet=ws;cells.forEach((row,i)=>row.forEach((c,j)=>{c.Address=()=>'A'+(i+1);}));
  return {cells,right,app:{ActiveWorkbook:wb,ActiveSheet:ws,Selection:range}};}
function run(app,id,answer=true){const window={Application:app,confirm:()=>answer};vm.runInNewContext(toolsCode(),{window,fetch:async()=>{throw Error('no network');},location:{pathname:'/addon/et/'},setTimeout,clearTimeout});return window.OfficeTools.run(id,{host:'et'});}
test('WPS 表格本地化工具：结果写到右侧，原数据不动；统一格式前先确认',async()=>{
  const s=sheet([[1234.5],['abc'],[undefined]]);assert.match(await run(s.app,'rmbUpper'),/1 个大写金额，1 个无法识别/);assert.equal(s.right[0][0].Value2,'壹仟贰佰叁拾肆元伍角');assert.equal(s.cells[0][0].Value2,1234.5);
  const base='11010519900307123',c='10X98765432'[[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2].reduce((a,x,i)=>a+x*Number(base[i]),0)%11];
  const id=sheet([[base+c],['bad']]);assert.match(await run(id.app,'idCardInfo'),/1 个号码有问题/);assert.equal(id.right[0][0].Value2,'1990-03-07');assert.equal(id.right[0][1].Value2,'男');
  const wk=sheet([['R8.11.12']]);await run(wk.app,'toSeireki');assert.equal(wk.right[0][0].Value2,46338);assert.equal(wk.right[0][0].NumberFormat,'yyyy/m/d');
  const z=sheet([['ＡＢ１']]);assert.match(await run(z.app,'zenHan',false),/已取消/);assert.equal(z.cells[0][0].Value2,'ＡＢ１');assert.match(await run(z.app,'zenHan'),/已统一 1/);assert.equal(z.cells[0][0].Value2,'AB1');
  const p=sheet([['1234567']]);await run(p.app,'postalFormat');assert.equal(p.cells[0][0].Value2,'123-4567');assert.equal(p.cells[0][0].NumberFormat,'@');
  const occ=sheet([[100]]);occ.right[0][0].Value2='旧';await assert.rejects(run(occ.app,'rmbUpper',false),/已取消/);assert.equal(occ.right[0][0].Value2,'旧');
  const f=sheet([['山田']]);await run(f.app,'furigana');assert.equal(f.right[0][0].Formula,'=PHONETIC(A1)');
  const two=sheet([['a','b']]);await assert.rejects(run(two.app,'idCardInfo'),/只选一列/);
});
