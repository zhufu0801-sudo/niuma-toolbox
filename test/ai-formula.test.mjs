import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { formulaPrompt, parseFormula } from '../lib/ai-tasks.mjs';
const src=fs.readFileSync(new URL('../public/ai-formula.js',import.meta.url),'utf8');
function sheet(){
  const cells={A1:'姓名',B1:'部门',C1:'工资',A2:'张三',B2:'销售',C2:5000,A3:'李四',B3:'行政',C3:4200,E2:'',F2:'已有'};
  const cell=a=>({get Value2(){return cells[a];},set Formula(f){cells[a]='#'+f;this.f=f;},get HasFormula(){return String(cells[a]||'').startsWith('#=');},Cells:{Count:1}});
  const used={Column:1,Row:1,Columns:{Count:3},Rows:{Count:3},Address:()=>'A1:C3',Cells:{Item:(i,j)=>cell('ABC'[j-1]+i)}};
  const ws={Name:'工资表',UsedRange:used,Range:a=>cell(a)};
  return {cells,app:{ActiveSheet:ws,ActiveCell:{Address:()=>'E2'}}};
}
function load(app,host='et'){
  const els={};const box={id:'',className:'',set innerHTML(h){for(const m of h.matchAll(/id="(af[A-Za-z]+)"/g))els[m[1]]={textContent:'',hidden:false,disabled:false,className:''};},querySelector:s=>els[s.slice(1)]};
  const fa={querySelector:()=>({parentNode:{insertBefore(){}},nextSibling:null})};
  const w={OfficeTools:{host:()=>host,app:()=>app,lib:{ask:async()=>true}},document:{createElement:()=>box,getElementById:id=>id==='formulaAssistant'?fa:null,addEventListener(){}}};
  w.window=w;vm.runInNewContext(src,w);return {w,els};
}
test('AI 帮写公式：WPS 表格只取表头和前 5 行示例，写入前检查结果单元格是空的',()=>{
  const {cells,app}=sheet();const {w,els}=load(app);assert.ok(els.afGo&&els.afApply);
  const ctx=w.AiFormula.wpsHost.context();
  assert.equal(JSON.stringify(ctx.headers),JSON.stringify([{col:'A',name:'姓名'},{col:'B',name:'部门'},{col:'C',name:'工资'}]));
  assert.equal(ctx.rows.length,2);assert.equal(ctx.target,'E2');assert.equal(ctx.sheet,'工资表');assert.equal(ctx.range,'A1:C3');
  assert.match(w.AiFormula.wpsHost.write('E2','=SUM($C$2:$C$3)','工资表'),/已写入 E2/);assert.equal(cells.E2,'#=SUM($C$2:$C$3)');
  assert.throws(()=>w.AiFormula.wpsHost.write('F2','=1','工资表'),/不是空的/);
  assert.throws(()=>w.AiFormula.wpsHost.write('G2','=1','别的表'),/已切换/);
});
test('AI 帮写公式：只在 WPS 表格里出现',()=>{const {w}=load({},'wps');assert.equal(w.AiFormula,undefined);});
test('AI 帮写公式：提示词只含表头和示例，公式经过安全检查',()=>{
  const m=formulaPrompt({intent:'按部门合计工资',context:{sheet:'工资表',range:'A1:C3',target:'E2',headers:[{col:'A',name:'姓名'},{col:'C',name:'工资'}],rows:[['张三',5000]]},lang:'ja'});
  assert.match(m[1].content,/A=姓名，C=工资/);assert.match(m[0].content,/日本語/);
  assert.equal(parseFormula('```json\n{"formula":"=SUMIF($B$2:$B$3,\\"销售\\",$C$2:$C$3)","explain":"按部门合计"}\n```').formula,'=SUMIF($B$2:$B$3,"销售",$C$2:$C$3)');
  assert.throws(()=>parseFormula('{"formula":"=WEBSERVICE(\\"http://x\\")"}'),/联网/);
  assert.throws(()=>parseFormula('{"formula":"=[a.xlsx]Sheet1!A1"}'),/其他文件/);
  assert.throws(()=>parseFormula('我不知道'),/没有给出公式/);
});
test('AI 帮写公式：Office 版读取表头和示例，只写入空白单元格',async()=>{
  const shim=fs.readFileSync(new URL('../public/office/office-shim.js',import.meta.url),'utf8');
  const cells={A1:'姓名',B1:'工资',A2:'张三',B2:5000,D2:''},formulas={};
  const L=n=>String.fromCharCode(65+n);
  const rng=(r,c,h,w)=>({load(){},get columnCount(){return w;},get rowCount(){return h;},get columnIndex(){return c;},get cellCount(){return h*w;},get address(){return '工资表!'+L(c)+(r+1);},
    getCell:(i,j)=>rng(r+i,c+j,1,1),getResizedRange:(dr,dc)=>rng(r,c,h+dr,w+dc),
    get values(){return Array.from({length:h},(_,i)=>Array.from({length:w},(_,j)=>{const a=L(c+j)+(r+i+1);return formulas[a]?42:(cells[a]??'');}));},
    get formulas(){return Array.from({length:h},(_,i)=>Array.from({length:w},(_,j)=>{const a=L(c+j)+(r+i+1);return formulas[a]||(cells[a]??'');}));},
    set formulas(x){formulas[L(c)+(r+1)]=x[0][0];}});
  const ws={name:'工资表',load(){},getUsedRange:()=>rng(0,0,2,2),getRange:a=>{const m=a.match(/^([A-Z])(\d+)$/);return rng(+m[2]-1,m[1].charCodeAt(0)-65,1,1);}};
  const w={Excel:{run:async f=>f({workbook:{worksheets:{getActiveWorksheet:()=>ws},getActiveCell:()=>rng(1,3,1,1)},sync:async()=>{}})}};w.window=w;
  vm.runInNewContext(shim,w);const H=w.AiFormulaHost;
  const ctx=await H.context();assert.equal(ctx.target,'D2');assert.equal(ctx.sheet,'工资表');assert.equal(ctx.headers.map(h=>h.col+h.name).join(),'A姓名,B工资');assert.equal(ctx.rows.length,1);
  assert.match(await H.write('D2','=SUM($B$2:$B$2)','工资表'),/结果：42/);assert.equal(formulas.D2,'=SUM($B$2:$B$2)');
  await assert.rejects(H.write('B2','=1','工资表'),/不是空的/);await assert.rejects(H.write('E2','=1','别的表'),/已切换/);
});
