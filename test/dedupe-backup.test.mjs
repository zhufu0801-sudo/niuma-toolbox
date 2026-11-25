// 删除重复行、一键备份。
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
import {createCoachServer} from '../server.mjs';
function bridge(a,host,fetchImpl,confirm=()=>true){const window={Application:a,confirm};vm.runInNewContext(toolsCode(),{window,fetch:fetchImpl,location:{pathname:'/addon/'+host+'/',search:''},URLSearchParams,Date,setTimeout,clearTimeout});return window.OfficeTools;}
function sheet(values,extra={}){
  const writes={};const range={Rows:{Count:values.length},Columns:{Count:values[0].length},MergeCells:false,HasFormula:false,Address:'A1:B9',Value2:values,
    Cells:{Item:(r,c)=>({Resize:(rows,cols)=>({set Value2(v){writes.top={r,rows,cols,v};},ClearContents(){writes.clear={r,rows,cols};}})})},...extra};
  return {a:{ActiveWorkbook:{Name:'s.xlsx'},ActiveSheet:{Name:'S'},Selection:range},writes};
}
test('删除重复行：保留第一次出现的行，按原顺序上移，下方多出的行清空',async()=>{
  const {a,writes}=sheet([['姓名','部门'],['小林','设计'],['小王','开发'],['小林','设计'],['=1+1','开发'],['=1+1','开发']]);
  const msg=await bridge(a,'et').run('removeDuplicates',{host:'et'});
  assert.equal(JSON.stringify(writes.top.v),JSON.stringify([['姓名','部门'],['小林','设计'],['小王','开发'],["'=1+1",'开发']]));
  assert.deepEqual([writes.clear.r,writes.clear.rows],[5,2]);assert.match(msg,/删除 2 行/);
  // 真机 0.6.2：表格里由加载项写入的内容按 Ctrl+Z 无法恢复，提示不能再承诺撤销。
  assert.doesNotMatch(msg,/可使用 WPS 撤销/);assert.match(msg,/撤销不能恢复/);
});
test('删除重复行：没有重复时不改动；有公式、合并单元格、只选一行时拒绝',async()=>{
  const none=sheet([['a',1],['b',2]]);assert.match(await bridge(none.a,'et').run('removeDuplicates',{host:'et'}),/没有重复/);assert.equal(none.writes.top,undefined);
  await assert.rejects(bridge(sheet([['a'],['a']],{HasFormula:null}).a,'et').run('removeDuplicates',{host:'et'}),/公式/);
  await assert.rejects(bridge(sheet([['a'],['a']],{MergeCells:null}).a,'et').run('removeDuplicates',{host:'et'}),/合并单元格/);
  await assert.rejects(bridge(sheet([['a']]).a,'et').run('removeDuplicates',{host:'et'}),/至少两行/);
  const cancel=sheet([['a'],['a']]);assert.equal(await bridge(cancel.a,'et',undefined,()=>false).run('removeDuplicates',{host:'et'}),'已取消。');assert.equal(cancel.writes.top,undefined);
});
test('备份：未保存过的文档提示先保存；有未保存修改时先确认',async()=>{
  const a={ActiveDocument:{Name:'a.docx',FullName:'a.docx',Path:''}};
  await assert.rejects(bridge(a,'wps').run('backupFile'),/先保存/);
  const b={ActiveDocument:{Name:'a.docx',FullName:'C:/x/a.docx',Path:'C:/x',Saved:false}};
  assert.equal(await bridge(b,'wps',async()=>{throw Error('不应请求');},()=>false).run('backupFile'),'已取消。');
});
test('备份接口：复制到备份文件夹、文件名带时间、同名不覆盖、只接受办公文件',async t=>{
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'bk-')),backupDir=path.join(dir,'备份'),src=path.join(dir,'报告.docx');await fs.writeFile(src,'原文件内容');
  const server=createCoachServer({backupDir});await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>server.close());
  const base='http://127.0.0.1:'+server.address().port,{token}=await (await fetch(base+'/api/session')).json();
  const call=(p,auth=token)=>fetch(base+'/api/backup',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':auth},body:JSON.stringify({path:p})});
  assert.equal((await call(src,'bad')).status,403);
  const one=await (await call(src)).json(),two=await (await call(src)).json();
  assert.match(path.basename(one.backupPath),/^报告-备份-\d{8}-\d{6}\.docx$/);assert.notEqual(one.backupPath,two.backupPath);
  assert.equal(await fs.readFile(one.backupPath,'utf8'),'原文件内容');assert.equal(await fs.readFile(src,'utf8'),'原文件内容');
  const exe=path.join(dir,'a.exe');await fs.writeFile(exe,'x');assert.equal((await call(exe)).status,422);
  assert.equal((await call('relative.docx')).status,422);assert.equal((await call(path.join(dir,'none.docx'))).status,422);
});
