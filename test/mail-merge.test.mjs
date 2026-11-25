// 按名单批量生成（2026-11-21）
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import os from 'node:os';import path from 'node:path';
import JSZip from 'jszip';import ExcelJS from 'exceljs';import {Document,Packer,Paragraph,TextRun,Header} from 'docx';
import {readList,inspectMerge,runMerge} from '../lib/mail-merge.mjs';
const dir=fs.mkdtempSync(path.join(os.tmpdir(),'mm-'));
async function makeList(rows,file='名单.xlsx'){const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('名单');rows.forEach(r=>ws.addRow(r));ws.getCell('C2').numFmt='yyyy"年"m"月"d"日"';const p=path.join(dir,file);await wb.xlsx.writeFile(p);return p;}
async function makeTemplate(file='邀请函.docx',header=false){
  // 占位符故意拆成几段文字（WPS / Word 常这样存）
  const doc=new Document({sections:[{headers:header?{default:new Header({children:[new Paragraph('致 {{姓名}}')]})}:undefined,children:[new Paragraph({children:[new TextRun('尊敬的 '),new TextRun('{{'),new TextRun({text:'姓',bold:true}),new TextRun('名}}'),new TextRun(' 先生/女士：')]}),new Paragraph('请于 {日期} 到 《部门》 报到。推荐阅读《三体》。{{职位}}')]}]});
  const p=path.join(dir,file);fs.writeFileSync(p,await Packer.toBuffer(doc));return p;}
const text=async(file,part='word/document.xml')=>(await (await JSZip.loadAsync(fs.readFileSync(file))).file(part).async('string')).replace(/<w:br w:type="page"\/>/g,'[分页]').replace(/<[^>]+>/g,'');
test('名单：第一行是列名，日期按 Excel 显示的样子取，空行跳过',async()=>{
  const list=await makeList([['姓名','部门','日期'],['张三','行政部',new Date(Date.UTC(2026,11,1))],[],['李四','财务部','2026/12/2']]);
  const {columns,rows}=await readList(list);assert.deepEqual(columns,['姓名','部门','日期']);assert.equal(rows.length,2);assert.equal(rows[0]['日期'],'2026年12月1日');
});
test('检查：用到的列、没用到的列、名单里没有的 {{占位符}}',async()=>{
  const list=await makeList([['姓名','部门','日期','电话'],['张三','行政部','12/1','1']]);const d=await inspectMerge({template:await makeTemplate(),list});
  assert.deepEqual(d.used,['姓名','部门','日期']);assert.deepEqual(d.unused,['电话']);assert.deepEqual(d.unknown,['职位']);assert.equal(d.count,1);
});
test('合成一个文档：每份之间分页；拆开的占位符也能替换；《三体》和名单没有的 {{职位}} 不动',async()=>{
  const list=await makeList([['姓名','部门','日期'],['张三','行政部','12月1日'],['李四 & 王','财务部','12月2日']]);const out=path.join(dir,'out1');
  const r=await runMerge({template:await makeTemplate(),list,outputDir:out});assert.equal(r.count,2);assert.match(path.basename(r.path),/邀请函-批量生成（2 份）\.docx/);
  const t=await text(r.path);assert.match(t,/尊敬的 张三 先生\/女士：请于 12月1日 到 行政部 报到。推荐阅读《三体》。\{\{职位\}\}\[分页\]尊敬的 李四 &amp; 王 先生/);
  const zip=await JSZip.loadAsync(fs.readFileSync(r.path));assert.ok((await zip.file('word/document.xml').async('string')).includes('李四 &amp; 王'),'特殊字符正确转义');
});
test('每行一个文件：文件名用指定的列，重名自动区分；页眉里的占位符也替换',async()=>{
  const list=await makeList([['姓名','部门','日期'],['张三','行政部','1'],['张三','财务部','2'],['','人事部','3']]);const out=path.join(dir,'out2');
  const r=await runMerge({template:await makeTemplate('通知.docx',true),list,mode:'separate',nameColumn:'姓名',outputDir:out});
  assert.equal(r.count,3);assert.deepEqual(fs.readdirSync(r.folder).sort(),['003.docx','张三-2.docx','张三.docx'].sort());
  assert.match(await text(path.join(r.folder,'张三-2.docx')),/财务部/);const z=await JSZip.loadAsync(fs.readFileSync(path.join(r.folder,'张三.docx')));const hdr=Object.keys(z.files).find(n=>/header\d*\.xml$/.test(n));assert.match(await text(path.join(r.folder,'张三.docx'),hdr),/致 张三/);
});
test('出错时说清楚：模板不是 docx、模板里没有占位符、名单只有表头',async()=>{
  const list=await makeList([['姓名'],['张三']]);await assert.rejects(runMerge({template:path.join(dir,'a.doc'),list,outputDir:dir}),/\.docx/);
  const plain=path.join(dir,'plain.docx');fs.writeFileSync(plain,await Packer.toBuffer(new Document({sections:[{children:[new Paragraph('没有占位符')]}]})));
  await assert.rejects(runMerge({template:plain,list,outputDir:dir}),/没有找到名单的列名.*\{\{姓名\}\}/);
  await assert.rejects(readList(await makeList([['姓名']],'空.xlsx')),/只有表头/);
});
test('名单日期：按格子的格式写（年月日、斜杠、带时间）',async()=>{
  const wb=new ExcelJS.Workbook(),ws=wb.addWorksheet('a');ws.addRow(['甲','乙','丙']);ws.addRow([new Date(Date.UTC(2026,0,5,9,7)),new Date(Date.UTC(2026,0,5)),new Date(Date.UTC(2026,0,5,14,30))]);
  ws.getCell('A2').numFmt='yyyy/m/d';ws.getCell('B2').numFmt='m"月"d"日"';ws.getCell('C2').numFmt='yyyy-mm-dd hh:mm';const p=path.join(dir,'日期.xlsx');await wb.xlsx.writeFile(p);
  const {rows}=await readList(p);assert.deepEqual(rows[0],{'甲':'2026/1/5','乙':'1月5日','丙':'2026-01-05 14:30'});
});
test('服务接口：检查和生成都要本机口令；文件放进“生成文件”文件夹；相对路径不收',async t=>{
  const {createCoachServer}=await import('../server.mjs');const d=fs.mkdtempSync(path.join(os.tmpdir(),'mms-'));t.after(()=>fs.rmSync(d,{recursive:true,force:true}));
  const server=createCoachServer({aiFile:path.join(d,'ai.json'),cardsFile:path.join(d,'c.json'),settingsFile:path.join(d,'s.json'),backupDir:path.join(d,'备份'),outputHome:d,privacyPluginDir:path.join(d,'plugins'),docToolsOptions:{platform:'linux',exec:async()=>({stdout:''})}});
  await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await (await fetch(base+'/api/session')).json()).token;
  const post=(body,k=token)=>fetch(base+'/api/tools/mail-merge',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':k},body:JSON.stringify(body)});
  const tpl=await makeTemplate('证书.docx'),list=await makeList([['姓名','部门','日期'],['王五','研发部','5日']],'证书名单.xlsx');
  assert.equal((await post({template:tpl,list,preview:true},'bad')).status,403);
  const pv=await (await post({template:tpl,list,preview:true})).json();assert.equal(pv.count,1);assert.deepEqual(pv.used,['姓名','部门','日期']);
  const bad=await post({template:'证书.docx',list});assert.equal(bad.status,422);assert.match((await bad.json()).error,/请先保存模板/);
  const r=await (await post({template:tpl,list})).json();assert.ok(r.path.startsWith(d),r.path);assert.match(await text(r.path),/王五/);
});
