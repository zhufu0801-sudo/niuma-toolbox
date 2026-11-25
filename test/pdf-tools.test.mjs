// PDF 工具（2026-11-21）
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import os from 'node:os';import path from 'node:path';import PDFDocument from 'pdfkit';import {createRequire} from 'node:module';
import {parsePages,mergePdfs,pickPages,splitPdf,imagesToPdf,pdfToImages,pdfInfo} from '../lib/pdf-tools.mjs';
const {PDFDocument:Lib}=createRequire(import.meta.url)('../lib/vendor/pdf-lib.min.cjs');
const dir=fs.mkdtempSync(path.join(os.tmpdir(),'pdft-')),out=path.join(dir,'生成文件');
const PNG=Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4500000000049454e44ae426082','hex');
function makePdf(name,n,size='A4'){return new Promise(r=>{const d=new PDFDocument({size,autoFirstPage:false}),c=[];d.on('data',x=>c.push(x));d.on('end',()=>{const p=path.join(dir,name);fs.writeFileSync(p,Buffer.concat(c));r(p);});for(let i=1;i<=n;i++){d.addPage();d.fontSize(40).text('Page '+i,100,100);}d.end();});}
const pagesOf=async f=>(await Lib.load(fs.readFileSync(f))).getPageCount();
test('页码：1-3,5,8- ；全角、顿号、～也认；超出范围说清楚',()=>{
  assert.deepEqual(parsePages('1-3,5,8-',10),[1,2,3,5,8,9,10]);assert.deepEqual(parsePages('２～４、１',10),[2,3,4,1]);assert.deepEqual(parsePages('-2',5),[1,2]);assert.deepEqual(parsePages('3-1',5),[3,2,1].sort());
  assert.throws(()=>parsePages('12',10),/只有 10 页/);assert.throws(()=>parsePages('a',10),/看不懂/);assert.throws(()=>parsePages('',10),/请填写页码/);
});
test('合并：按顺序接起来，文件名用第一个文件；原文件不动',async()=>{
  const a=await makePdf('合同.pdf',2),b=await makePdf('附件.pdf',3),before=fs.readFileSync(a);
  const r=await mergePdfs({paths:[a,b],outputDir:out});assert.equal(r.pages,5);assert.match(path.basename(r.path),/^合同-合并（2 个文件）\.pdf$/);assert.equal(await pagesOf(r.path),5);assert.ok(before.equals(fs.readFileSync(a)));
  assert.match(path.basename((await mergePdfs({paths:[a,b],outputDir:out})).path),/-2\.pdf$/,'重名加编号，不覆盖');
  await assert.rejects(mergePdfs({paths:[a],outputDir:out}),/至少选择 2 个/);await assert.rejects(mergePdfs({paths:[a,'附件.pdf'],outputDir:out}),/完整位置/);
  fs.writeFileSync(path.join(dir,'坏.pdf'),'not a pdf');await assert.rejects(mergePdfs({paths:[a,path.join(dir,'坏.pdf')],outputDir:out}),/打不开/);
});
test('提取页面 / 删除页面',async()=>{
  const a=await makePdf('报告.pdf',6);const k=await pickPages({path:a,pages:'1-2,5',outputDir:out});assert.equal(k.pages,3);assert.match(path.basename(k.path),/报告-第1-2,5页\.pdf/);
  const d=await pickPages({path:a,pages:'2',mode:'remove',outputDir:out});assert.equal(d.pages,5);assert.match(path.basename(d.path),/报告-删除第2页\.pdf/);assert.equal(await pagesOf(d.path),5);
  await assert.rejects(pickPages({path:a,pages:'1-',mode:'remove',outputDir:out}),/没有页面/);
});
test('拆分：每页一个、每 N 页一个、按段拆',async()=>{
  const a=await makePdf('名册.pdf',5);const r=await splitPdf({path:a,every:1,outputDir:out});assert.equal(r.count,5);assert.equal(r.files[0],'01-第1页.pdf');
  const r2=await splitPdf({path:a,every:2,outputDir:out});assert.deepEqual(r2.files,['01-第1-2页.pdf','02-第3-4页.pdf','03-第5页.pdf']);assert.match(r2.folder,/名册-拆分-2$/);
  const r3=await splitPdf({path:a,ranges:'1-2，3-',outputDir:out});assert.deepEqual(r3.files,['01-第1-2页.pdf','02-第3-5页.pdf']);assert.equal(await pagesOf(path.join(r3.folder,r3.files[1])),3);
  await assert.rejects(splitPdf({path:a,every:5,outputDir:out}),/一个和原文件一样/);
});
test('图片转 PDF：A4 每张一页，横图横放；不支持的格式说清楚',async()=>{
  const p=path.join(dir,'扫描1.png');fs.writeFileSync(p,PNG);const j=path.join(dir,'扫描2.jpg');fs.writeFileSync(j,PNG);
  const r=await imagesToPdf({paths:[p,p],outputDir:out});assert.equal(r.pages,2);assert.match(path.basename(r.path),/扫描1-等 2 张图片\.pdf/);const doc=await Lib.load(fs.readFileSync(r.path));assert.equal(Math.round(doc.getPage(0).getWidth()),595);
  const f=await imagesToPdf({paths:[j],size:'fit',outputDir:out});assert.equal(await pagesOf(f.path),1);
  fs.writeFileSync(path.join(dir,'x.png'),'nope');await assert.rejects(imagesToPdf({paths:[path.join(dir,'x.png')],outputDir:out}),/只支持 JPG、PNG/);await assert.rejects(imagesToPdf({paths:[path.join(dir,'a.gif')],outputDir:out}),/不是 JPG 或 PNG/);
});
test('PDF 转图片：每页一张 PNG（用 pdf.js 画），可以只转部分页',async()=>{
  const a=await makePdf('通知.pdf',3);const r=await pdfToImages({path:a,pages:'1,3',outputDir:out});assert.equal(r.count,2);assert.deepEqual(r.files,['通知-01.png','通知-03.png']);
  const png=fs.readFileSync(path.join(r.folder,r.files[0]));assert.equal(png.subarray(1,4).toString(),'PNG');assert.ok(png.readUInt32BE(16)>1200,'150 dpi 下 A4 宽约 1240 像素');
  assert.equal((await pdfInfo({path:a})).pages,3);
});
test('服务接口：要本机口令；结果进“生成文件”；不认识的动作返回 404',async t=>{
  const {createCoachServer}=await import('../server.mjs');const d=fs.mkdtempSync(path.join(os.tmpdir(),'pdfs-'));t.after(()=>fs.rmSync(d,{recursive:true,force:true}));const revealed=[];
  const server=createCoachServer({aiFile:path.join(d,'ai.json'),cardsFile:path.join(d,'c.json'),settingsFile:path.join(d,'s.json'),backupDir:path.join(d,'备份'),outputHome:d,privacyPluginDir:path.join(d,'plugins'),docToolsOptions:{platform:'linux',exec:async(cmd,args)=>{revealed.push([cmd,args]);return {stdout:''};}}});
  await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await (await fetch(base+'/api/session')).json()).token;
  const post=(body,k=token)=>fetch(base+'/api/tools/pdf',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':k},body:JSON.stringify(body)});
  const a=await makePdf('甲.pdf',2),b=await makePdf('乙.pdf',1);
  assert.equal((await post({action:'merge',paths:[a,b]},'x')).status,403);
  assert.equal((await (await post({action:'info',path:a})).json()).pages,2);
  const r=await (await post({action:'merge',paths:[a,b]})).json();assert.ok(r.path.startsWith(d),r.path);assert.equal(r.pages,3);
  const s=await (await post({action:'split',path:a,every:1})).json();assert.equal(s.count,2);
  const bad=await post({action:'pick',path:a,pages:'9'});assert.equal(bad.status,422);assert.match((await bad.json()).error,/只有 2 页/);
  assert.equal((await post({action:'nope'})).status,404);
});
