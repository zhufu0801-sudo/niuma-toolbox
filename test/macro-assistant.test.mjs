import test from 'node:test';import assert from 'node:assert/strict';import vm from 'node:vm';import fs from 'node:fs';
const ctx={window:{}};vm.createContext(ctx);for(const f of ['formula-engine.js','macro-engine.js'])vm.runInContext(fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8'),ctx);const M=ctx.window.OfficeMacro;
test('四类宏仅处理工作表副本，代码为 ASCII，中文路径不受 VBA 导入编码影响',()=>{
 for(const kind of Object.keys(M.types)){const g=M.generate({kind,range:"'测试表'!A1:C10",book:'C:\\测试\\工资.xlsx'});assert.match(g.code,/src.Copy After/);assert.match(g.code,/Set area = dst.Range/);assert.ok(!/[^\x00-\x7f]/.test(g.code));assert.ok(g.code.split('\r\n').every(l=>l.length<1000));assert.ok(!g.code.includes('Shell('));assert.ok(!g.code.includes('Kill '));}
 const g=M.generate({kind:'sort',range:'表!A1:C10',book:'C:\\'+'长'.repeat(300)+'.xlsx',column:2,descending:true});assert.match(g.code,/xlDescending/);assert.ok(g.code.split('\r\n').every(l=>l.length<1000));
});
test('宏识别未知需求不猜测；安装只新增模块不执行',()=>{
 assert.equal(M.infer('按金额降序排序'),'sort');assert.equal(M.infer('删除重复行'),'dedupe');assert.equal(M.infer('任意脚本'),null);let code='',ran=false;
 const g=M.generate({kind:'trim',range:'表!A1:C10',book:'test.xlsx'});const app={ActiveWorkbook:{FullName:'test.xlsx',VBProject:{VBComponents:{Add(){return {Name:'',CodeModule:{AddFromString(s){code=s;}}};}}}},Run(){ran=true;}};
 M.install(app,g);assert.equal(code,g.code);assert.equal(ran,false);assert.throws(()=>M.install({ActiveWorkbook:{FullName:'other.xlsx'}},g),/切回/);assert.throws(()=>M.install({ActiveWorkbook:{FullName:'test.xlsx'}},g),/VBA/);
});
test('宏写入失败清理本次新模块',()=>{let removed=false;const g=M.generate({kind:'format',range:'表!A1:C10',book:'test.xlsx'});assert.throws(()=>M.install({ActiveWorkbook:{FullName:'test.xlsx',VBProject:{VBComponents:{Add(){return {CodeModule:{AddFromString(){throw Error('denied');}}};},Remove(){removed=true;}}}}},g),/安装宏未完成/);assert.equal(removed,true);});
