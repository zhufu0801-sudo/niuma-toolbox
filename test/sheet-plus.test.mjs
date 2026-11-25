import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
import {toolsCode} from './helpers/tools-code.mjs';
import {makeBook,makeApp} from './helpers/et-mock.mjs';
const code=toolsCode();
function load(app,answer=true){const asked=[];const window={Application:app,confirm:m=>{asked.push(m);return typeof answer==='function'?answer(m):answer;}};vm.runInNewContext(code,{window,fetch:async()=>{throw Error('no network');},location:{pathname:'/addon/et/'},setTimeout,clearTimeout,URLSearchParams});const T=window.OfficeTools;T.asked=asked;return T;}
function setup(sheets,sel,activeCell){const book=makeBook(sheets),app=makeApp(book);app.select(sel,activeCell);return {book,app,ws:book.list[0]};}
const val=(ws,r,c)=>ws.cell(r,c).Value2;
const plain=x=>JSON.parse(JSON.stringify(x));

test('文本转数字：数字、千分位、全角、百分比、日期转真值；0 开头和超过 15 位保留',async()=>{
  const {app,ws}=setup({S:[['1,234'],['１２'],['50%'],['2026/1/5'],['0012'],['110101199001011234'],['(30)'],['abc']]},'A1:A8');
  const T=load(app),r=await T.run('textToNumber',{host:'et'});
  assert.match(r,/已转换 5 个.*日期 1 个/);
  assert.deepEqual([1,2,3,4,5,6,7,8].map(i=>val(ws,i,1)),[1234,12,0.5,46027,'0012','110101199001011234',-30,'abc']);
  assert.equal(ws.cell(4,1).NumberFormat,'yyyy/m/d');assert.match(T.asked[0],/2 个以 0 开头或超过 15 位/);
});
test('提取数字 / 中文 / 英文 / 去括号：写到右边空白列，右边有内容时拒绝',async()=>{
  let s=setup({S:[['单价12.5元，数量3'],['Apple苹果（红）']]},'A1:A2');let T=load(s.app);
  assert.match(await T.run('extractNumbers',{host:'et'}),/B1:B2/);assert.deepEqual([val(s.ws,1,2),val(s.ws,2,2)],['12.5 3','']);
  s=setup({S:[['Apple苹果（红）','']]},'A1:A1');T=load(s.app);await T.run('extractChinese',{host:'et'});assert.equal(val(s.ws,1,2),'苹果红');
  s=setup({S:[['Apple Pie苹果']]},'A1');T=load(s.app);await T.run('extractLetters',{host:'et'});assert.equal(val(s.ws,1,2),'Apple Pie');
  s=setup({S:[['产品A（旧款）[停产]']]},'A1');T=load(s.app);await T.run('removeBrackets',{host:'et'});assert.equal(val(s.ws,1,2),'产品A');
  s=setup({S:[['a1','已有']]},'A1');T=load(s.app);await assert.rejects(T.run('extractNumbers',{host:'et'}),/已有内容/);assert.equal(val(s.ws,1,2),'已有');
});
test('两表核对：同样大小逐格标黄；大小不同按第一列对照',async()=>{
  let s=setup({S:[['张三',100,'','张三',100],['李四',200,'','李四',250]]},['A1:B2','D1:E2']);let T=load(s.app);
  assert.match(await T.run('compareRanges',{host:'et'}),/已标出 1 处/);assert.equal(s.ws.cell(2,2).Interior.Color,s.ws.cell(2,5).Interior.Color);assert.notEqual(s.ws.cell(2,2).Interior.Color,16777215);assert.equal(s.ws.cell(1,1).Interior.Color,16777215);
  s=setup({S:[['张三',1,'','张三',1],['李四',2,'','王五',3],['赵六',4,'','','']]},['A1:B3','D1:E2']);T=load(s.app);
  assert.match(await T.run('compareRanges',{host:'et'}),/只在第一块 2 行、只在第二块 1 行/);
  s=setup({S:[['a']]},'A1');T=load(s.app);await assert.rejects(T.run('compareRanges',{host:'et'}),/按住 Ctrl/);
});
test('工资条：每人一行带表头，中间空一行，原表不变',async()=>{
  const s=setup({工资:[['姓名','工资'],['张三',5000],['李四',6000]]},'A1:B3');const T=load(s.app);
  assert.match(await T.run('payrollSlips',{host:'et'}),/2 人/);const out=s.book.list[1];assert.equal(out.Name,'工资条');
  assert.deepEqual([[1,1],[2,1],[3,1],[4,1],[5,1]].map(([r,c])=>val(out,r,c)??''),['姓名','张三','','姓名','李四']);assert.equal(val(s.ws,2,1),'张三');
});
test('合并单元格编号：大小不一的合并单元格依次 1、2、3',async()=>{
  const s=setup({S:[['x'],[''],[''],[''],[''],['']]},'A1:A6');s.ws.merges.push({r1:1,c1:1,r2:3,c2:1},{r1:4,c1:1,r2:4,c2:1},{r1:5,c1:1,r2:6,c2:1});const T=load(s.app);
  assert.match(await T.run('mergedNumbering',{host:'et'}),/1～3/);assert.deepEqual([1,4,5].map(r=>val(s.ws,r,1)),[1,2,3]);assert.match(T.asked[0],/1 个已有内容会被覆盖/);
});
test('排名不打乱顺序：右边写 RANK 公式',async()=>{
  const s=setup({S:[[90],[80],[95]]},'A1:A3');const T=load(s.app);await T.run('rankKeepOrder',{host:'et'});
  assert.equal(s.ws.cell(1,2).Formula,'=IF(ISNUMBER(A1),RANK(A1,$A$1:$A$3,0),"")');assert.equal(s.ws.cell(3,2).Formula,'=IF(ISNUMBER(A3),RANK(A3,$A$1:$A$3,0),"")');
});
test('多列合一格（可选分隔符）、一格拆多行、多列堆成一列',async()=>{
  let s=setup({S:[['北京','朝阳',''],['上海','','浦东']]},'A1:C2');let T=load(s.app);
  await T.run('joinColumns',{host:'et',separator:'-'});assert.deepEqual([val(s.ws,1,4),val(s.ws,2,4)],['北京-朝阳','上海-浦东']);
  s=setup({S:[['项目','人员'],['A','张三、李四'],['B','王五']]},'B2:B3');T=load(s.app);assert.match(await T.run('splitToRows',{host:'et'}),/1 行被拆开，共 3 行/);
  const out=s.book.list[1];assert.deepEqual([1,2,3,4].map(r=>[val(out,r,1),val(out,r,2)]),[['项目','人员'],['A','张三'],['A','李四'],['B','王五']]);
  s=setup({S:[['人员','到期'],['张三、李四',46296]]},'A2:A2');s.ws.cell(2,2).NumberFormat='yyyy/m/d';T=load(s.app);await T.run('splitToRows',{host:'et'});
  const dated=s.book.list[1];assert.deepEqual([val(dated,2,2),dated.cell(2,2).NumberFormat,dated.cell(3,2).NumberFormat,dated.cell(1,2).NumberFormat],[46296,'yyyy/m/d','yyyy/m/d','General'],'日期格式带到新表');
  s=setup({S:[[1,4],[2,''],[3,6]]},'A1:B3');T=load(s.app);await T.run('stackColumns',{host:'et'});const st=s.book.list[1];assert.deepEqual([1,2,3,4,5].map(r=>val(st,r,1)),[1,2,3,4,6]);
});
test('删除空白行：只删整行都空的行（看整张表的列）',async()=>{
  const s=setup({S:[['a',1],['',''],['','备注'],['b',2]]},'A1:A4');const T=load(s.app);assert.match(await T.run('deleteBlankRows',{host:'et'}),/已删除 1 个/);
  assert.deepEqual([1,2,3].map(r=>[val(s.ws,r,1)??'',val(s.ws,r,2)??'']),[['a',1],['','备注'],['b',2]]);
});
test('工作表目录、批量改名（不重名、可加前后缀）',async()=>{
  let s=setup({一月:[['x']],二月:[['y']]},'A1');let T=load(s.app);await T.run('sheetIndex',{host:'et'});const idx=s.book.list[0];assert.equal(idx.Name,'目录');assert.match(idx.cell(2,1).Formula,/HYPERLINK\("#'一月'!A1","一月"\)/);
  s=setup({'2025一月':[['x']],'2025二月':[['y']]},'A1');T=load(s.app);assert.match(await T.run('renameSheets',{host:'et',find:'2025',replace:'2026',suffix:'表'}),/已改名 2 个/);assert.deepEqual(s.book.list.map(w=>w.Name),['2026一月表','2026二月表']);
  s=setup({'A':[['x']],'B':[['y']]},'A1');T=load(s.app);await assert.rejects(T.run('renameSheets',{host:'et',find:'A',replace:'B'}),/重名/);
  await assert.rejects(T.run('renameSheets',{host:'et'}),/请填写/);
});
test('按列拆分成多表、多表合并（表头不同的跳过）',async()=>{
  let s=setup({数据:[['部门','姓名'],['销售','甲'],['财务','乙'],['销售','丙']]},'A1:B4','A2');let T=load(s.app);
  assert.match(await T.run('splitByColumn',{host:'et'}),/2 张工作表/);assert.deepEqual(s.book.list.map(w=>w.Name),['数据','销售','财务']);assert.equal(val(s.book.list[1],3,2),'丙');
  s=setup({一:[['名','额'],['a',1]],二:[['名','额'],['b',2]],三:[['别的']]},'A1');T=load(s.app);assert.match(await T.run('mergeSheets',{host:'et'}),/已合并 2 张.*跳过表头不同的：三/);
  const m=s.book.list[3];assert.deepEqual([1,2,3].map(r=>[val(m,r,1),val(m,r,3)]),[['名','来源工作表'],['a','一'],['b','二']]);
});
test('两列互换、行列转置、加前后缀、IFERROR、颜色统计',async()=>{
  let s=setup({S:[['a',1],['b',2]]},'A1:B2');let T=load(s.app);await T.run('swapColumns',{host:'et'});assert.deepEqual([val(s.ws,1,1),val(s.ws,1,2)],[1,'a']);
  s=setup({S:[[1,2,3],[4,5,6]]},'A1:C2');T=load(s.app);await T.run('transposeRange',{host:'et'});const t=s.book.list[1];assert.deepEqual([[1,1],[1,2],[3,1],[3,2]].map(([r,c])=>val(t,r,c)),[1,4,3,6]);
  s=setup({S:[['001'],[5],['']]},'A1:A3');T=load(s.app);assert.match(await T.run('addPrefixSuffix',{host:'et',prefix:'BJ-'}),/已处理 2 个/);assert.deepEqual([val(s.ws,1,1),val(s.ws,2,1)],['BJ-001','BJ-5']);
  s=setup({S:[['=VLOOKUP(A1,B:C,2,0)'],['=IFERROR(A1,"")'],['x']]},'A1:A3');T=load(s.app);assert.match(await T.run('wrapIferror',{host:'et'}),/已处理 1 个/);assert.equal(s.ws.cell(1,1).Formula,'=IFERROR(VLOOKUP(A1,B:C,2,0),"")');
  s=setup({S:[[10],[20],[5]]},'A1:A3');s.ws.cell(1,1).Interior.Color=255;s.ws.cell(1,1).Interior.ColorIndex=3;s.ws.cell(2,1).Interior.Color=255;s.ws.cell(2,1).Interior.ColorIndex=3;T=load(s.app);
  const r=await T.run('colorSumCount',{host:'et'});assert.match(r,/#FF0000：2 个，数字合计 30/);assert.match(r,/无填充：1 个，数字合计 5/);
});
test('隔行填色、到期提醒：再点一次取消；聚光灯开关',async()=>{
  let s=setup({S:[['h'],[1],[2],[3]]},'A1:A4');let T=load(s.app);await T.run('bandRows',{host:'et'});assert.equal(s.ws.cf.length,1);assert.equal(s.ws.cf[0].formula,'=MOD(ROW()-2,2)=1');assert.match(await T.run('bandRows',{host:'et'}),/已取消/);assert.equal(s.ws.cf.length,0);
  s=setup({S:[[46000],[46100]]},'A1:A2');T=load(s.app);await T.run('dueReminder',{host:'et'});assert.equal(s.ws.cf.length,2);assert.match(await T.run('dueReminder',{host:'et'}),/已取消/);assert.equal(s.ws.cf.length,0);
  s=setup({S:[['a']]},'A1');T=load(s.app);assert.match(await T.run('spotlight',{host:'et'}),/不能跟着选区自动移动/);assert.equal(s.ws.cf[0].formula,'=OR(ROW()=1,COLUMN()=1)');assert.match(await T.run('spotlight',{host:'et'}),/已关闭/);assert.equal(s.ws.cf.length,0);
  const ev={list:[],AddApiEventListener(n,f){this.list.push([n,f]);},RemoveApiEventListener(n,f){this.list=this.list.filter(x=>x[1]!==f);}};
  s=setup({S:[['a','b'],['c','d']]},'A1');s.app.ApiEvent=ev;T=load(s.app);assert.match(await T.run('spotlight',{host:'et'}),/跟着选中的单元格移动/);assert.equal(ev.list[0][0],'SheetSelectionChange');
  s.app.select('B2');ev.list[0][1]();assert.equal(s.ws.cf.length,1);assert.equal(s.ws.cf[0].formula,'=OR(ROW()=2,COLUMN()=2)');await T.run('spotlight',{host:'et'});assert.equal(ev.list.length,0);assert.equal(s.ws.cf.length,0);
});
test('打印表头、一页宽、一键图表',async()=>{
  const s=setup({S:[['月份','销量'],['1月',10],['2月',20]]},'A1');const T=load(s.app);await T.run('printTitles',{host:'et'});assert.equal(s.ws.PageSetup.PrintTitleRows,'$1:$1');
  await T.run('fitOnePageWide',{host:'et'});assert.equal(s.ws.PageSetup.FitToPagesWide,1);assert.equal(s.ws.PageSetup.FitToPagesTall,false);
  s.app.select('A1:B3');assert.match(await T.run('quickChart',{host:'et'}),/柱形图/);assert.equal(s.ws.charts[0].type,51);assert.equal(s.ws.charts[0].source,'$A$1:$B$3');
});
test('表格体检：分类规则',()=>{
  const T=load(setup({S:[['x']]},'A1').app),C=T.SheetCheck.classify;
  assert.deepEqual(plain(C({value:'1,200'})),['textNumber']);assert.deepEqual(plain(C({value:' 张三'})),['spaces']);assert.deepEqual(plain(C({value:'张三 '})),['hidden']);
  assert.deepEqual(plain(C({value:'=SUM(A1:A3)'})),['formulaText']);assert.deepEqual(plain(C({value:0,formula:'=VLOOKUP(A1,B:C,2)',text:'5'})),['approx']);
  assert.deepEqual(plain(C({value:0,formula:'=VLOOKUP(A1,B:C,2,FALSE)',text:'#N/A'})),['na']);assert.deepEqual(plain(C({value:110101199001011000,text:'1.10101E+17'})),['longNumber','scientific']);
  assert.deepEqual(plain(C({value:'110101199001011234'})),['longText']);assert.deepEqual(plain(C({value:'正常'})),[]);
  assert.deepEqual(plain(T.SheetCheck.split('A1,IF(B1>0,"a,b",C1),2')),['A1','IF(B1>0,"a,b",C1)','2']);
});
test('表格体检：报告文字、手动计算提示、不修改内容',async()=>{
  const s=setup({S:[['1,200','=VLOOKUP(A1,B:C,2)'],[' 张三','ok']]},'A1');s.app.Calculation=-4135;const T=load(s.app);
  const r=await T.run('sheetCheck',{host:'et'});assert.match(r,/发现 4 类问题/);assert.match(r,/看起来是数字、实际是文字：1 处（A1）/);assert.match(r,/近似匹配/);assert.match(r,/计算方式是“手动”/);assert.match(r,/没有修改任何内容/);assert.equal(val(s.ws,1,1),'1,200');
});
test('表格整理工具只在 WPS 表格登记，带输入框的工具有 fields',()=>{
  const T=load(setup({S:[['x']]},'A1').app),ids=T.list('et').filter(x=>x.group==='sheetplus').map(x=>x.id);
  assert.ok(ids.length>=32,ids.length);assert.equal(T.list('wps').some(x=>x.group==='sheetplus'),false);
  assert.deepEqual(plain(T.list('et').find(x=>x.id==='addPrefixSuffix').fields.map(f=>f.key)),['prefix','suffix']);
  const xml=fs.readFileSync(new URL('../public/ribbon.xml',import.meta.url),'utf8');for(const id of ids)assert.ok(xml.includes('id="'+id+'"'),id);
});
test('打码写右边一列、带单位求和只报结果、一列整理成表、细长表分栏',async()=>{
  let s=setup({S:[['110101199003071234'],['13812345678'],['张三']]},'A1:A3'),T=load(s.app);
  assert.match(await T.run('maskData',{host:'et'}),/2 个/);assert.deepEqual([1,2,3].map(r=>val(s.ws,r,2)),['110101********1234','138****5678','']);assert.equal(val(s.ws,1,1),'110101199003071234');
  s=setup({S:[['12元','3元'],['无',5]]},'A1:B2');T=load(s.app);assert.match(await T.run('unitSum',{host:'et'}),/合计：20元（3 个单元格，1 个没有数字已跳过）/);
  s=setup({S:[['张三'],['138'],['北京'],['李四'],['139'],['上海']]},'A1:A6');T=load(s.app);assert.match(await T.run('listToTable',{host:'et',perRow:3}),/2 行、每行 3 列/);
  const t=s.book.list[1];assert.deepEqual([1,2].map(r=>[val(t,r,1),val(t,r,2),val(t,r,3)]),[['张三','138','北京'],['李四','139','上海']]);
  await assert.rejects(load(setup({S:[['a']]},'A1').app).run('listToTable',{host:'et'}),/每条记录有几行/);
  s=setup({S:[['名','分'],['a',1],['b',2],['c',3],['d',4],['e',5],['f',6]]},'A1:B7');T=load(s.app);assert.match(await T.run('columnsForPrint',{host:'et',groups:2}),/分成 2 栏（每栏 3 行）/);
  const p=s.book.list[1];assert.deepEqual([val(p,1,1),val(p,1,4),val(p,2,4),val(p,4,5)],['名','名','d',6]);
});
