// 日期工具（2026-11-21）：日期识别、格式、中国调休 / 日本祝日、工作日、各宿主工具。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';

const ctx={};ctx.window=ctx;vm.runInNewContext(fs.readFileSync(new URL('../public/date-engine.js',import.meta.url),'utf8'),ctx);const D=ctx.OfficeDates;
const ymd=d=>d&&D.ymd(d);

test('日期识别：各种写法（全角、和暦、元年、星期、8 位数字）；不存在的日子返回空',()=>{
  const cases={'2026/11/21':'2026-11-21','2026-1-5':'2026-01-05','2026.11.21':'2026-11-21','2026年11月21日':'2026-11-21','２０２６年１１月２１日（土）':'2026-11-21','2026年11月21日 星期六':'2026-11-21','20261121':'2026-11-21','令和8年11月21日':'2026-11-21','令和元年5月1日':'2019-05-01','R8.11.21':'2026-11-21','平成31年4月30日':'2019-04-30','11月21日':'2026-11-21'};
  for(const [s,v] of Object.entries(cases))assert.equal(ymd(D.parse(s,{defaultYear:2026})),v,s);
  for(const s of ['2026/2/30','2026/13/1','abc','12345','11/21'])assert.equal(D.parse(s),null,s);
  assert.equal(ymd(D.parse(46347,{allowSerial:true})),'2026-11-21');assert.equal(D.toSerial(D.parse('2026/11/21')),46347);
});
test('日期格式：7 种写法；令和元年；R 简写',()=>{
  const d=D.parse('2026/11/21');
  assert.deepEqual(Array.from(D.FORMATS,f=>D.format(d,f.id)),['2026/11/21','2026-11-21','2026年11月21日','2026年11月21日 星期六','2026年11月21日（土）','令和8年11月21日','R8.11.21']);
  assert.equal(D.format(D.parse('2019/5/1'),'wareki'),'令和元年5月1日');assert.equal(D.format(D.parse('2019/4/30'),'wareki'),'平成31年4月30日');
});
test('日本の祝日：2026 年（振替休日・国民の休日を含む）と 2019・2020 年の特例',()=>{
  assert.deepEqual(Object.keys(D.jpHolidays(2026)).sort(),['2026-01-01','2026-01-12','2026-02-11','2026-02-23','2026-03-20','2026-04-29','2026-05-03','2026-05-04','2026-05-05','2026-05-06','2026-07-20','2026-08-11','2026-09-21','2026-09-22','2026-09-23','2026-10-12','2026-11-03','2026-11-23']);
  const h19=D.jpHolidays(2019);for(const k of ['2019-04-30','2019-05-01','2019-05-02','2019-10-22'])assert.ok(h19[k],k);
  const h20=D.jpHolidays(2020);assert.equal(h20['2020-07-24'],'スポーツの日');assert.equal(h20['2020-08-10'],'山の日');assert.ok(!h20['2020-10-12']);
});
test('中国：放假和调休（国务院办公厅 2025–2027 年通知）；调休的周末算工作日',()=>{
  assert.equal(D.dayInfo(D.parse('2026/2/14'),'cn').work,true,'春节前周六调休上班');assert.equal(D.dayInfo(D.parse('2026/2/16'),'cn').name,'春节');
  assert.equal(D.dayInfo(D.parse('2025/10/8'),'cn').work,false);assert.equal(D.dayInfo(D.parse('2027/9/26'),'cn').work,true);assert.equal(D.dayInfo(D.parse('2027/6/9'),'cn').name,'端午节');
  const k=D.countWorkdays(D.parse('2026/2/1'),D.parse('2026/2/28'),'cn');assert.equal(k.workdays,16);assert.equal(k.makeup,2);
  assert.deepEqual(Array.from(D.countWorkdays(D.parse('2028/1/1'),D.parse('2028/1/10'),'cn').missingYears),[2028]);
  assert.equal(ymd(D.addWorkdays(D.parse('2026/9/30'),3,'cn')),'2026-10-10','国庆后第 3 个工作日是调休的周六');
  assert.equal(ymd(D.addWorkdays(D.parse('2026/12/25'),5,'jp',{yearEnd:true})),'2027-01-07');assert.equal(ymd(D.addWorkdays(D.parse('2026/1/9'),-1,'jp')),'2026-01-08');
});

function bridge(a,host,confirm=()=>true){const asked=[];const window={Application:a,confirm:m=>{asked.push(m);return confirm(m);}};vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/'+host+'/',search:''},URLSearchParams,Date,setTimeout,clearTimeout});return {T:window.OfficeTools,asked};}
function sheet(values,formats={}){const cellsMap={};const cell=(i,j)=>{const k=i+','+j;return cellsMap[k]||(cellsMap[k]={Value2:values[i-1]?.[j-1],NumberFormat:formats[k]||'General'});};
  const range={Rows:{Count:values.length},Columns:{Count:values[0].length},Cells:{Item:cell},Address:'A1'};const sheets=[];
  const wb={Name:'s.xlsx',FullName:'C:/x/s.xlsx',Worksheets:{get Count(){return 1+sheets.length;},Item:()=>({}),Add(){const data={};const ws={Name:'',Cells:{Item:(r,c)=>({r,c})},Range:(a,b)=>({set Value2(v){ws.data=v;},Font:{}}),Columns:{Item:()=>({})}};sheets.push(ws);return ws;}}};
  return {a:{ActiveWorkbook:wb,ActiveSheet:{Name:'S'},Selection:range,ActiveCell:cell(1,1)},cell,sheets};}
test('统一日期格式（表格）：认出的日期写成真正的日期 + 格式，和暦写成文字；不是日期的格子不动；只确认一次',async()=>{
  const {a,cell}=sheet([['2026/11/21','令和8年1月5日'],['20261201','备注'],[46347,123]],{'3,1':'yyyy/m/d'});
  const {T,asked}=bridge(a,'et');const msg=await T.runFromUi('unifyDates',{format:'iso'});
  assert.equal(asked.length,1);assert.match(asked[0],/认出 4 个日期.*2 个格子不是日期/s);assert.match(msg,/已统一 4 个日期/);
  assert.equal(cell(1,1).Value2,46347);assert.equal(cell(1,1).NumberFormat,'yyyy-mm-dd');assert.equal(cell(1,2).Value2,46027);assert.equal(cell(2,1).Value2,46357);
  assert.equal(cell(2,2).Value2,'备注');assert.equal(cell(3,2).Value2,123,'普通数字不是日期');
  const w=sheet([['2026/11/21']]);await bridge(w.a,'et').T.runFromUi('unifyDates',{format:'wareki'});assert.equal(w.cell(1,1).Value2,'令和8年11月21日');assert.equal(w.cell(1,1).NumberFormat,'@');
});
function wordDoc(text,sel){const doc={text};const range=(s,e)=>{const r={Start:s,End:e,get Text(){return doc.text.slice(r.Start,r.End);},set Text(v){doc.text=doc.text.slice(0,r.Start)+v+doc.text.slice(r.End);r.End=r.Start+v.length;},Select(){sel=[r.Start,r.End];},get Duplicate(){return range(r.Start,r.End);},Collapse(d){if(d===0)r.Start=r.End;else r.End=r.Start;}};return r;};
  return {a:{ActiveDocument:{Name:'a.docx',ReadOnly:false,Range:range},Selection:{get Start(){return sel[0];},get End(){return sel[1];},get Text(){return doc.text.slice(sel[0],sel[1]);},get Range(){return range(sel[0],sel[1]);},Collapse(){sel=[sel[1],sel[1]];}}},doc,get sel(){return sel;}};}
test('统一日期格式（文字）：只改选区里的日期，版本号、电话不动；可撤销',async()=>{
  const t='会议定于2026/11/21（土）召开，截止 令和8年12月1日，版本 1.2.3，电话 03-1234-5678。';const w=wordDoc(t,[0,t.length]);
  const msg=await bridge(w.a,'wps').T.runFromUi('unifyDates',{format:'cn'});
  assert.equal(w.doc.text,'会议定于2026年11月21日召开，截止 2026年12月1日，版本 1.2.3，电话 03-1234-5678。');assert.match(msg,/2 处/);
});
test('插入今天日期：文字插在光标处；表格写入真正的日期；格式可选',async()=>{
  const now=new Date(),d=new Date(now.getFullYear(),now.getMonth(),now.getDate());
  const w=wordDoc('甲乙',[1,1]);await bridge(w.a,'wps').T.runFromUi('insertToday',{format:'wareki'});assert.equal(w.doc.text,'甲'+D.format(d,'wareki')+'乙');
  const s=sheet([['']]);await bridge(s.a,'et').T.runFromUi('insertToday',{format:'slash'});assert.equal(s.cell(1,1).Value2,D.toSerial(d));assert.equal(s.cell(1,1).NumberFormat,'yyyy/mm/dd');
});
test('算工作日：两天之间、几个工作日后；没有数据的年份会提示；不改文档、不弹确认',async()=>{
  const {T,asked}=bridge({ActiveDocument:{Name:'a.docx'},Selection:{}},'wps');
  let msg=await T.runFromUi('workdayCalc',{start:'2026/2/1',end:'2026/2/28',region:'cn'});assert.match(msg,/共 16 个工作日/);assert.match(msg,/调休上班 2 天/);
  msg=await T.runFromUi('workdayCalc',{start:'2026/9/30',days:'3',region:'cn'});assert.match(msg,/2026年10月10日 星期六/);
  msg=await T.runFromUi('workdayCalc',{start:'2026/12/25',days:'5',region:'jpYearEnd'});assert.match(msg,/2027年1月7日（木）/);
  msg=await T.runFromUi('workdayCalc',{start:'2028/1/3',end:'2028/1/7',region:'cn'});assert.match(msg,/没有 2028 年的放假安排/);
  assert.equal(asked.length,0);
});
test('按月汇总：自动找日期列和金额列，表头保留，结果放新工作表',async()=>{
  const s=sheet([['日期','客户','金额'],['2026/1/5','甲',100],['2026/1/20','乙',50],['看不懂','丁',9]],{'2,1':'General'});
  const msg=await bridge(s.a,'et').T.runFromUi('dateSummary',{period:'month'});
  assert.deepEqual(JSON.parse(JSON.stringify(s.sheets[0].data)),[['期间','金额（合计）','笔数'],['2026年1月',150,2],['总计',150,2]]);assert.match(msg,/1 行日期看不懂/);
  const q=sheet([['2026/1/5',1],['2026/4/1',2],['2026/5/1',3]]);await bridge(q.a,'et').T.runFromUi('dateSummary',{period:'quarter'});
  assert.deepEqual(JSON.parse(JSON.stringify(q.sheets[0].data)).slice(1),[['2026年第1季度',1,1],['2026年第2季度',5,2],['总计',6,3]]);
});

test('押印欄（文字）：插在光标处，2 行表格，靠右，盖章格约 2 厘米；职位可自定义',async()=>{
  const made=[];const mkCell=()=>({Range:{Text:'',ParagraphFormat:{},Font:{}}});
  const d={Name:'a.docx',ReadOnly:false,Tables:{Add(r,rows,cols){const cs=Array.from({length:rows},()=>Array.from({length:cols},mkCell));const t={rows,cols,Rows:{Item:i=>(t['row'+i]??={}),Alignment:0},Columns:{Item:j=>(t['col'+j]??={})},Borders:{},Cell:(i,j)=>cs[i-1][j-1],cs};made.push(t);return t;}}};
  const r={Start:3,End:3};const a={ActiveDocument:d,Selection:{Range:r}};
  const msg=await bridge(a,'wps').T.runFromUi('approvalBox',{labels:'担当、課長、部長'});
  const t=made[0];assert.equal(t.rows,2);assert.equal(t.cols,3);assert.deepEqual(t.cs[0].map(c=>c.Range.Text),['担当','課長','部長']);assert.deepEqual(t.cs[1].map(c=>c.Range.Text),['','','']);
  assert.equal(t.Rows.Alignment,2);assert.equal(t.row2.Height,57);assert.equal(t.col1.Width,57);assert.ok(t.Borders.Enable);assert.match(msg,/担当 \/ 課長 \/ 部長/);
  await assert.rejects(bridge(a,'wps').T.run('approvalBox',{host:'wps',labels:'一,二,三,四,五,六,七,八,九'}),/最多 8 个/);
});
test('押印欄（表格）：从选中格子开始写，有内容时不覆盖',async()=>{
  let written=null,counta=0;const area={set Value2(v){written=v;},Borders:{}};
  const ws={Range:()=>area,Cells:{Item:(r,c)=>({r,c})},Rows:{Item:()=>({})},Columns:{Item:()=>({})}};
  const sel={Rows:{Count:1},Columns:{Count:1},Cells:{Item:()=>({Row:3,Column:2})},Address:'B3'};
  const a={ActiveWorkbook:{Name:'s.xlsx',FullName:'C:/s.xlsx'},ActiveSheet:Object.assign(ws,{Name:'S'}),Selection:sel,WorksheetFunction:{CountA:()=>counta}};
  await bridge(a,'et').T.runFromUi('approvalBox',{labels:'经办,审核,批准'});assert.deepEqual(JSON.parse(JSON.stringify(written)),[['经办','审核','批准'],['','','']]);
  counta=1;await assert.rejects(bridge(a,'et').T.run('approvalBox',{host:'et'}),/已经有内容/);
});
