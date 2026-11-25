import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const read=f=>fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8');
// 极简 Excel 模拟：单元格存 value/formula，支持 getRange('A1:B3')、getCell、copy、dataValidation、autoFilter、worksheets.add
function col(s){let n=0;for(const ch of s)n=n*26+ch.charCodeAt(0)-64;return n;}
function parse(a){const m=String(a).replace(/\$/g,'').replace(/^.*!/,'').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);const r1=+m[2]-1,c1=col(m[1])-1,r2=m[3]?+m[4]-1:r1,c2=m[3]?col(m[3])-1:c1;return {r1,c1,r2,c2};}
function letter(n){let s='';n++;while(n){const k=(n-1)%26;s=String.fromCharCode(65+k)+s;n=Math.floor((n-1)/26);}return s;}
function makeBook(){const book={name:'book.xlsx',sheets:[],selection:null};
 function sheet(name,cells={}){const ws={name,cells,validations:[],filter:null,protection:{protected:false,load(){}},load(){},
  getRange(a){return range(ws,parse(a));},getRangeByIndexes(r,c,h,w){return range(ws,{r1:r,c1:c,r2:r+h-1,c2:c+w-1});},
  copy(){const n=sheet(name+' (2)',JSON.parse(JSON.stringify(cells)));book.sheets.push(n);return n;},activate(){},
  autoFilter:{enabled:false,load(){},apply(r,ci,crit){ws.filter={ci,crit};},remove(){}}};return ws;}
 function range(ws,b){const k=(r,c)=>r+','+c;const get=(r,c)=>ws.cells[k(r,c)]||{v:'',f:''};const R={worksheet:ws,load(){},
  get rowCount(){return b.r2-b.r1+1;},get columnCount(){return b.c2-b.c1+1;},get rowIndex(){return b.r1;},get columnIndex(){return b.c1;},
  get address(){return ws.name+'!'+letter(b.c1)+(b.r1+1)+(b.r2>b.r1||b.c2>b.c1?':'+letter(b.c2)+(b.r2+1):'');},
  get values(){const o=[];for(let r=b.r1;r<=b.r2;r++){const l=[];for(let c=b.c1;c<=b.c2;c++)l.push(get(r,c).v);o.push(l);}return o;},
  set values(x){x.forEach((l,i)=>l.forEach((v,j)=>{ws.cells[k(b.r1+i,b.c1+j)]={v:typeof v==='string'&&v[0]==="'"?v.slice(1):v,f:''};}));},
  get formulas(){const o=[];for(let r=b.r1;r<=b.r2;r++){const l=[];for(let c=b.c1;c<=b.c2;c++){const x=get(r,c);l.push(x.f||x.v);}o.push(l);}return o;},
  set formulas(x){x.forEach((l,i)=>l.forEach((f,j)=>{ws.cells[k(b.r1+i,b.c1+j)]={v:0,f};}));},
  get text(){return this.values.map(l=>l.map(String));},
  set numberFormat(x){R._nf=x;},
  getCell(i,j){return range(ws,{r1:b.r1+i,c1:b.c1+j,r2:b.r1+i,c2:b.c1+j});},
  getResizedRange(dr,dc){return range(ws,{...b,r2:b.r2+dr,c2:b.c2+dc});},
  getMergedAreasOrNullObject(){return {isNullObject:true,load(){}};},
  clear(){for(let r=b.r1;r<=b.r2;r++)for(let c=b.c1;c<=b.c2;c++)delete ws.cells[k(r,c)];},
  merge(){ws.merged=(ws.merged||[]).concat([R.address]);},unmerge(){},select(){book.selection=R;},
  format:{font:{},autofitColumns(){}},
  dataValidation:{clear(){ws.validations=[];},set rule(r){ws.validations.push(r);},set errorAlert(a){ws.alert=a;},ignoreBlanks:true}};return R;}
 const ctx={workbook:{get name(){return book.name;},load(){},getSelectedRange:()=>book.selection,
  worksheets:{getItem:n=>book.sheets.find(s=>s.name===n),getItemOrNullObject:n=>book.sheets.find(s=>s.name===n)||{isNullObject:true,load(){}},add(){const s=sheet('Sheet'+(book.sheets.length+1));book.sheets.push(s);return s;},getActiveWorksheet:()=>book.sheets[0]}},sync:async()=>{}};
 book.sheet=sheet;book.ctx=ctx;return book;}
function load(book){const w={localStorage:{getItem:()=>null,setItem(){}}};w.window=w;w.Excel={run:fn=>fn(book.ctx)};
 for(const f of ['quick-engine.js','formula-engine.js','quick-host.js','office/office-tools.js','office/office-shim.js'])vm.runInNewContext(read(f),w);return w;}
function data(){const b=makeBook();const ws=b.sheet('数据',{});b.sheets.push(ws);
 const rows=[['部门','姓名','金额'],['销售','甲',10],['销售','乙',20],['财务','丙',5]];rows.forEach((l,i)=>l.forEach((v,j)=>{ws.cells[i+','+j]={v,f:''};}));return {b,ws};}
test('Office 版技巧导航：读取选区、预览、在副本中执行，原表不动',async()=>{
 const {b,ws}=data();const w=load(b);b.selection=ws.getRange('A1:C4');
 const p=await w.QuickHost.capture();assert.equal(p.sheet,'数据');assert.equal(p.address,'A1:C4');assert.equal(p.rows[1][2],10);
 const plan=w.QuickHost.plan('groupSummary',p,{column:1,valueColumn:3});assert.equal(plan.count,2);
 const msg=await w.QuickHost.apply(plan);assert.match(msg,/2 类汇总/);const out=b.sheets[1];assert.deepEqual(out.getRange('A1:C3').values,[['部门','次数','合计'],['销售',2,30],['财务',1,5]]);
 b.selection=ws.getRange('C2:C4');const p2=await w.QuickHost.capture();const pl2=w.QuickHost.plan('clearConstants',p2,{});const m2=await w.QuickHost.apply(pl2);assert.match(m2,/处理副本「数据 \(2\)」/);
 assert.equal(ws.getRange('C2').values[0][0],10,'原表不变');const copy=b.sheets.find(s=>s.name==='数据 (2)');assert.equal(copy.getRange('C2').values[0][0],'');
 b.selection=ws.getRange('B2:B4');const p3=await w.QuickHost.capture();const m3=await w.QuickHost.apply(w.QuickHost.plan('preventDuplicates',p3,{}));assert.match(m3,/处理副本/);
 b.selection=ws.getRange('C2:C4');const p4=await w.QuickHost.capture();ws.cells['1,2']={v:99,f:''};await assert.rejects(w.QuickHost.apply(w.QuickHost.plan('clearConstants',p4,{})),/来源数据已改变/);
});
test('Office 版中文公式助手：预览、只写空白格、撤回',async()=>{
 const {b,ws}=data();const w=load(b);const H=w.OfficeFormulaHost;
 b.selection=ws.getRange('C2:C4');const src=await H.capture();b.selection=ws.getRange('E2');const dst=await H.capture();
 const plan=await H.preview({},{kind:'round',sheet:'数据',source:src.reference,target:dst.reference,digits:'0',book:src.book});assert.equal(plan.plan.formulas.length,3);
 const rec=await H.apply({},plan);assert.match(ws.getRange('E2').formulas[0][0],/ROUND\('数据'!\$C2,0\)/);
 assert.equal(await H.undo({},rec),3);assert.equal(ws.getRange('E2').formulas[0][0],'');
 ws.cells['1,4']={v:'有内容',f:''};await assert.rejects(H.preview({},{kind:'round',sheet:'数据',source:src.reference,target:dst.reference,digits:'0'}),/已有内容/);
});
test('Office 版：只适用于 WPS 的操作指引换成 Office 入口，并有日文',()=>{
 const w={localStorage:{getItem:()=>null,setItem(){}}};w.window=w;
 for(const f of ['i18n-ja.js','i18n-ja-extra.js','method-cards.js','quick-engine.js','formula-engine.js','quick-host.js','quick-catalog.js','office/office-tools.js','office/office-shim.js'])vm.runInNewContext(read(f),w);
 const all=[...w.QuickCatalog.items,...w.MethodCards.cards].filter(c=>c.kind==='guide');
 const left=all.filter(c=>/WPS/.test(c.label+c.desc+c.steps));assert.equal(left.map(c=>c.id).join(),'');
 for(const k of Object.keys(w.OfficeShimGuides.overrides))for(const v of Object.values(w.OfficeShimGuides.overrides[k]))assert.ok(w.OfficeI18nJa[v]||w.OfficeI18nJa[v]==='',k+' 缺日文：'+v.slice(0,20));
});
test('Word 多行 OCR / AI 文字按原顺序插入独立段落，空行保留',async()=>{
 const w=load(makeBook()),inserted=[];const paragraph={insertParagraph(text,where){assert.equal(where,'After');inserted.push(text);return paragraph;}};
 const selection={text:'原文',load(){},insertParagraph:paragraph.insertParagraph,insertText(text){inserted.push(text);}};
 w.Word={run:fn=>fn({document:{getSelection:()=>selection},sync:async()=>{}})};w.OfficeShim.setHost('word');
 await w.AiTextHost.insert('第一行\n\n日本語\r\n最后一行',false);assert.deepEqual(inserted,['第一行','','日本語','最后一行']);
 inserted.length=0;selection.text='';await w.AiTextHost.insert('甲\n乙',false);assert.deepEqual(inserted,['甲\r乙']);
});
