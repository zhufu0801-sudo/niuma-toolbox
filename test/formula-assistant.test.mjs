import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const context={window:{}};vm.createContext(context);
for(const file of ['formula-engine.js','formula-host.js'])vm.runInContext(fs.readFileSync(new URL('../public/'+file,import.meta.url),'utf8'),context);
const F=context.window.OfficeFormula,H=context.window.OfficeFormulaHost;
const base={kind:'lookup',sheet:'员工表',source:'员工表!A2:A4',table:'工资表!A2:D100',target:'员工表!E2',keyColumn:1,returnColumn:4};
test('公式引用中文工作表、按行推进查找键，锁定查找表且不吞掉非匹配错误',()=>{
 const p=F.plan(base);assert.equal(p.formulas.length,3);assert.match(p.formulas[0],/IFNA\(VLOOKUP/);assert.match(p.formulas[0],/'员工表'!\$A2/);assert.match(p.formulas[2],/'员工表'!\$A4/);assert.match(p.formulas[2],/'工资表'!\$A\$2:\$D\$100,4,FALSE/);assert.match(p.formulas[0],/SUBSTITUTE/);assert.equal(p.target.r2,4);
});
test('左向查找用 INDEX MATCH；XLOOKUP 使用精确匹配；单引号转义',()=>{
 assert.match(F.plan({...base,keyColumn:4,returnColumn:1}).formulas[0],/INDEX\('工资表'!\$A\$2:\$A\$100,MATCH/);
 assert.match(F.plan({...base,kind:'xlookup'}).formulas[0],/XLOOKUP\('员工表'!\$A2/);
 assert.equal(F.full(F.ref("'O''Brien 表'!$B$2:C3",'默认')), "'O''Brien 表'!$B$2:$C$3");
});
test('中文需求识别与参数提取，不认识的需求明确返回空',()=>{
 for(const [text,kind] of [['跨表查找工资','lookup'],['XLOOKUP 匹配编号','xlookup'],['按部门合计金额','sumif'],['统计姓名出现的次数','countif'],['判断成绩大于等于 60','if'],['取编号前4位','left'],['提取年份','year']])assert.equal(F.infer(text).kind,kind,text);
 assert.equal(F.infer('四舍五入保留3位小数').options.digits,3);assert.equal(F.infer('请分析业务').kind,null);assert.equal(F.infer('判断大于等于60').options.operator,'>=');
});
test('拒绝循环引用、非法引用、跨文件注入与错误行列',()=>{
 for(const patch of [{target:'员工表!A3'},{target:'工资表!D50'},{source:'A1:A1048577'},{source:'XFE1'},{source:'[外部.xlsx]表!A1'},{source:'A1)+WEBSERVICE("x")'},{source:'A2:B3'},{target:'E2:F4'},{target:'E2:E3'},{keyColumn:8}])assert.throws(()=>F.plan({...base,...patch}));
 assert.throws(()=>F.plan({...base,kind:'concat',second:'B2:B3'}));
});
test('文字参数编码为公式字符串，不作为代码拼接',()=>{
 const p=F.plan({...base,fallback:'");WEBSERVICE("https://x")'});assert.match(p.formulas[0],/""\);WEBSERVICE\(""https:\/\/x""\)"/);
 assert.throws(()=>F.plan({...base,kind:'if',operator:'>1,WEBSERVICE('}));
 const c=F.plan({...base,kind:'countif'});assert.match(c.formulas[0],/COUNTIF/);assert.match(c.formulas[0],/"="&SUBSTITUTE/);
});
test('汇总生成一个结果，日期与条件处理保留空白',()=>{
 assert.equal(F.plan({...base,kind:'sum',source:'A2:D20'}).formulas.length,1);
 assert.match(F.plan({...base,kind:'datediff',second:'B2:B4'}).formulas[0],/INT\('员工表'!\$B2\)-INT\('员工表'!\$A2\)/);
 for(const kind of Object.keys(F.types))assert.ok(F.plan({...base,kind,second:'B2:B4'}).formulas[0].startsWith('='),kind);
});
function mock(){
 const cells=new Map();let writes=0,failAt=0;
 const book={FullName:'C:\\独立测试.xlsx',ReadOnly:false,Worksheets:{Item(name){if(!['员工表','工资表'].includes(name))throw Error('missing sheet');return {Name:name,ProtectContents:false,Range(address){const r=F.ref(address,name);return {Calculate(){},Cells:{Item(i,j){const key=name+'!'+(r.r1+i-1)+','+(r.c1+j-1);if(!cells.has(key)){let formula='';const c={Value2:'',MergeCells:false,HasArray:false,Text:'',get HasFormula(){return formula!=='';},get Formula(){return formula;},set Formula(v){writes++;if(writes===failAt)throw Error('模拟写入失败');formula=v;},ClearContents(){formula='';this.Value2='';}};cells.set(key,c);}return cells.get(key);}}};}};}}};
 return {app:{ActiveWorkbook:book},cells,fail(n){failAt=n;},at(row){return book.Worksheets.Item('员工表').Range('E'+row).Cells.Item(1,1);}};
}
test('真实宿主接口路径：写入并精确撤回；保持来源与格式',()=>{
 const m=mock(),p=H.preview(m.app,base),record=H.apply(m.app,p);assert.equal(m.at(2).Formula,p.plan.formulas[0]);assert.equal(H.undo(m.app,record),3);assert.equal(m.at(2).Formula,'');assert.equal(m.at(4).Formula,'');
});
test('预览后目标被修改 / 工作簿切换均停止；不覆盖原有值',()=>{
 const m=mock(),p=H.preview(m.app,base);m.at(3).Value2='已有数据';assert.throws(()=>H.apply(m.app,p),/已有内容/);assert.equal(m.at(2).Formula,'');m.at(3).Value2='';m.app.ActiveWorkbook.FullName='C:\\其他.xlsx';assert.throws(()=>H.apply(m.app,p),/切换/);
});
test('中途写入失败回滚；撤回遇到用户修改保留所有结果',()=>{
 const m=mock();m.fail(2);assert.throws(()=>H.apply(m.app,H.preview(m.app,base)),/已移除/);assert.equal(m.at(2).Formula,'');assert.equal(m.at(3).Formula,'');
 const n=mock(),record=H.apply(n.app,H.preview(n.app,base));n.at(3).Formula='=100';assert.throws(()=>H.undo(n.app,record),/已被修改/);assert.ok(n.at(2).Formula);assert.equal(n.at(3).Formula,'=100');
});
test('WPS 省略工作表引号后仍可写入撤回，文字常量变化仍被保护',()=>{
 function normalized(){const m=mock();for(let row=2;row<=4;row++){const c=m.at(row),descriptor=Object.getOwnPropertyDescriptor(c,'Formula');Object.defineProperty(c,'Formula',{get(){return descriptor.get.call(c).split(/("(?:[^"]|"")*")/).map((part,i)=>i%2?part:part.replace(/'员工表'!/g,'员工表!').replace(/'工资表'!/g,'工资表!')).join('');},set:descriptor.set});}return m;}
 const m=normalized(),record=H.apply(m.app,H.preview(m.app,base));assert.equal(H.undo(m.app,record),3);
 const n=normalized();n.fail(2);assert.throws(()=>H.apply(n.app,H.preview(n.app,base)),/已移除/);assert.equal(n.at(2).Formula,'');
 const a=normalized(),r=H.apply(a.app,H.preview(a.app,{...base,fallback:"'员工表'!"}));a.at(3).Formula=a.at(3).Formula.replace('"\'员工表\'!"','"员工表!"');assert.throws(()=>H.undo(a.app,r),/已被修改/);
});
test('日期转星期：识别“星期/周几”，空白保留，结果用 TEXT 转成星期几',()=>{
 assert.equal(F.infer('把日期变成星期几').kind,'weekday');assert.equal(F.infer('日付を曜日に').kind,'weekday');
 const p=F.plan({...base,kind:'weekday'});assert.equal(p.formulas[0],"=IF('员工表'!$A2=\"\",\"\",TEXT('员工表'!$A2,\"aaaa\"))");assert.ok(p.warnings.some(w=>/aaaa/.test(w)));
});
