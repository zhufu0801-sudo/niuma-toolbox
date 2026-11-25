import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const read=f=>fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8');
const plain=x=>JSON.parse(JSON.stringify(x));
function ctx(files){const c={};c.window=c;vm.runInNewContext(files.map(read).join('\n'),c);return c;}
const S=ctx(['sheet-logic.js']).OfficeSheetLogic;
test('文字转数字的判断：日期、百分比、负数、全角；编号和长号码保留',()=>{
  assert.deepEqual(plain(S.parseText('2026年1月5日')),{v:46027,date:true});assert.equal(S.parseText('2026/2/30'),null);
  assert.deepEqual(plain(S.parseText('１２．５')),{v:12.5,pct:false});assert.equal(S.parseText('(1,000)').v,-1000);assert.equal(S.parseText('￥99').v,99);
  assert.equal(S.parseText('007'),null);assert.equal(S.parseText('6222021234567890123'),null);assert.equal(S.parseText('12a'),null);
});
test('两表核对的对照规则、拆行、分组、改名、合并',()=>{
  assert.deepEqual(plain(S.compare([[1,'a'],[2,'b']],[[1,'a'],[2,'c']]).cells),[[1,1]]);
  const c=S.compare([['甲',1],['乙',2],['丙',3]],[['甲',1],['乙',9]]);assert.deepEqual(plain([c.onlyA,c.onlyB,c.diff]),[[2],[],[[1,1]]]);
  assert.equal(S.compare([[' 甲 ']],[['甲']]).cells.length,0);
  assert.deepEqual(plain(S.splitRows(['h'],[['a；b']],0).rows),[['h'],['a'],['b']]);
  const g=S.groupBy([['部门'],['销售'],[''],['销售']],0);assert.deepEqual(plain(g.order),['销售','（空白）']);
  assert.deepEqual(plain(S.renamePlan(['1月','2月'],{suffix:'表'})),[['1月','1月表'],['2月','2月表']]);assert.throws(()=>S.renamePlan(['A','B'],{find:'A',replace:'B'}),/重名/);
  assert.equal(S.cleanName('a/b:c'),'a_b_c');assert.equal(S.cleanName(''),'空白');
  const m=S.mergePlan([{name:'一',values:[['名'],['a']]},{name:'二',values:[['名'],['b'],['']]},{name:'三',values:[['别']]}]);assert.equal(m.parts,2);assert.deepEqual(plain(m.rows),[['名','来源工作表'],['a','一'],['b','二']]);assert.deepEqual(plain(m.mismatch),['三']);
});
test('条件格式规则的识别只认工具箱自己的',()=>{
  assert.ok(S.isSpot('=OR(ROW()=5,COLUMN()=3)'));assert.ok(!S.isSpot('=OR(ROW()=5,COLUMN()=3,A1>0)'));
  assert.ok(S.isBand('=MOD(ROW()-2,2)=1'));assert.ok(!S.isBand('=MOD(ROW(),2)=1'));
  assert.ok(S.isDue('=AND(ISNUMBER(A2),A2<TODAY())'));assert.ok(S.isDue('=AND(ISNUMBER(A2),A2>=TODAY(),A2<=TODAY()+30)'));assert.ok(!S.isDue('=A2<TODAY()'));
});
test('常用公式：多条件查找、一对多、多条件求和、去重计数、排名、年龄、反向查找',()=>{
  const F=ctx(['formula-engine.js']).OfficeFormula;
  let p=F.plan({kind:'multilookup',sheet:'S',source:'A2:A3',second:'B2:B3',target:'D2',table:'T!A2:D100',keyColumn:1,key2Column:2,returnColumn:4});assert.equal(p.formulas[0],"=IF(OR('S'!$A2=\"\",'S'!$B2=\"\"),\"\",IFERROR(INDEX('T'!$D$2:$D$100,MATCH(1,INDEX(('T'!$A$2:$A$100='S'!$A2)*('T'!$B$2:$B$100='S'!$B2),0),0)),\"未找到\"))");
  p=F.plan({kind:'onetomany',sheet:'S',source:'A2',target:'C2',table:'T!A2:B9',keyColumn:1,returnColumn:2});assert.match(p.formulas[0],/TEXTJOIN\("、",TRUE,IF\('T'!\$A\$2:\$A\$9='S'!\$A2,'T'!\$B\$2:\$B\$9,""\)\)/);
  p=F.plan({kind:'sumifs',sheet:'S',source:'A2',second:'B2',target:'C2',table:'T!A2:C9',keyColumn:1,key2Column:2,returnColumn:3});assert.match(p.formulas[0],/SUMIFS\('T'!\$C\$2:\$C\$9,'T'!\$A\$2:\$A\$9,'S'!\$A2,'T'!\$B\$2:\$B\$9,'S'!\$B2\)/);
  p=F.plan({kind:'distinctcount',sheet:'S',source:'A2:A50',target:'C1'});assert.equal(p.formulas.length,1);assert.match(p.formulas[0],/SUMPRODUCT\(\('S'!\$A\$2:\$A\$50<>""\)\/COUNTIF/);
  p=F.plan({kind:'rank',sheet:'S',source:'B2:B4',target:'C2'});assert.equal(p.formulas.length,3);assert.match(p.formulas[2],/RANK\('S'!\$B4,'S'!\$B\$2:\$B\$4,0\)/);
  p=F.plan({kind:'age',sheet:'S',source:'B2',target:'C2'});assert.match(p.formulas[0],/DATEDIF\('S'!\$B2,TODAY\(\),"Y"\)/);
  p=F.plan({kind:'reverse',sheet:'S',source:'A2',target:'C2',table:'T!A2:D9',keyColumn:3,returnColumn:4});assert.match(p.formulas[0],/INDEX\('T'!\$D\$2:\$D\$9,MATCH/);
  assert.equal(F.infer('按部门和姓名两个条件查找工资').kind,'multilookup');assert.equal(F.infer('统计不重复的客户个数').kind,'distinctcount');assert.equal(F.infer('算员工工龄').kind,'age');assert.equal(F.infer('按销量排名').kind,'rank');
});
// ---- Excel（Office.js）模拟 ----
function excel(values,formulas){
  const writes=[],state={values:values.map(r=>r.slice())};
  const cell=(i,j)=>({set values(v){state.values[i][j]=v[0][0];writes.push(['v',i,j,v[0][0]]);},set numberFormat(v){writes.push(['f',i,j,v[0][0]]);},set formulas(v){writes.push(['F',i,j,v[0][0]]);},format:{fill:{set color(c){writes.push(['c',i,j,c]);}}}});
  const sheets=[{name:'S'}];
  const sel={load(){},get values(){return state.values;},formulas:formulas||values.map(r=>r.map(v=>v)),numberFormat:values.map(r=>r.map(()=>'General')),rowCount:values.length,columnCount:values[0].length,rowIndex:0,columnIndex:0,address:'S!A1',getCell:cell,worksheet:{getRangeByIndexes:(r,c,h,w)=>({load(){},formulas:Array.from({length:h},()=>Array(w).fill('')),address:'S!B1',set values(v){writes.push(['out',r,c,v]);}})}};
  const context={workbook:{getSelectedRange:()=>sel,worksheets:{load(){},items:sheets}},sync:async()=>{}};
  return {api:{Excel:{run:fn=>fn(context)}},writes,state};
}
test('Excel 版：文本转数字先确认、提取写右边、工具清单与 WPS 版一致',async()=>{
  const c=ctx(['sheet-logic.js','office/office-sheetplus.js']),X=c.OfficeSheetPlus.tools;
  let m=excel([['1,234'],['0012'],['abc']]);const first=await X.textToNumber(m.api,{});assert.match(first.confirm,/将把 1 个.*1 个以 0 开头/);assert.equal(m.writes.length,0);
  assert.match(await X.textToNumber(m.api,{confirmed:true}),/已转换 1 个/);assert.deepEqual(plain(m.writes.find(w=>w[0]==='v')),['v',0,0,1234]);
  m=excel([['单价12元']]);assert.match(await X.extractNumbers(m.api,{}),/右边一列/);assert.deepEqual(plain(m.writes[0][3]),[['12']]);
  await assert.rejects(X.exportSheets(),/移动或复制/);
  const wpsIds=['spotlight','sheetCheck','compareRanges','textToNumber','bandRows','deleteBlankRows','extractNumbers','extractChinese','extractLetters','removeBrackets','addPrefixSuffix','alignNames','wrapIferror','mergedNumbering','rankKeepOrder','joinColumns','splitToRows','stackColumns','swapColumns','transposeRange','sheetIndex','renameSheets','splitByColumn','mergeSheets','exportSheets','printTitles','fitOnePageWide','quickChart','payrollSlips','dueReminder','progressBar','colorSumCount','autoBorder','maskData','unitSum','columnsForPrint','listToTable','rangeToImage','picturesByName'];
  assert.deepEqual(plain(c.OfficeSheetPlus.list.map(x=>x[0])).sort(),wpsIds.slice().sort());for(const id of wpsIds)assert.equal(typeof X[id],'function',id);
});
test('打码、带单位求和、分栏、一列整理成表',()=>{
  assert.equal(S.mask('110101199003071234'),'110101********1234');assert.equal(S.mask('11010119900307123X'),'110101********123X');assert.equal(S.mask('13812345678'),'138****5678');
  assert.equal(S.mask('090-1234-5678'),'090-****-5678');assert.equal(S.mask('6222020200112233445'),'6222 **** **** 3445');assert.equal(S.mask('tanaka@example.co.jp'),'t***@example.co.jp');assert.equal(S.mask('张三'),null);assert.equal(S.mask(''),null);
  assert.deepEqual(plain(S.unitSum(['12元','3.5元','1,200元',8,'无',''])),{sum:1223.5,count:4,skipped:1,units:['元']});assert.equal(S.unitSum(['0.1kg','0.2kg']).sum,0.3);
  assert.deepEqual(plain(S.columnsFor([['名','分'],['a',1],['b',2],['c',3]],2)),{rows:[['名','分','','名','分'],['a',1,'','c',3],['b',2,'','','']],per:2});
  assert.deepEqual(plain(S.listToTable(['张三','138','北京','','李四','139'],3)),{rows:[['张三','138','北京'],['李四','139','']],rest:2});
});
test('Excel 版：打码写右边、带单位求和、一列整理成表、存图片和插图片给出手动步骤',async()=>{
  const c=ctx(['sheet-logic.js','office/office-sheetplus.js']),X=c.OfficeSheetPlus.tools;
  let m=excel([['13812345678'],['备注']]);assert.match(await X.maskData(m.api,{}),/1 个/);
  m=excel([['12元'],['8元']]);assert.match(await X.unitSum(m.api,{}),/合计：20元（2 个单元格）/);
  await assert.rejects(X.listToTable(excel([['a'],['b']]).api,{}),/每条记录有几行/);
  await assert.rejects(X.rangeToImage(),/复制为图片/);await assert.rejects(X.picturesByName(),/插入 → 图片/);
});
