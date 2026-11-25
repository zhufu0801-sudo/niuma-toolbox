import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import PDFDocument from 'pdfkit';
import {createDocTools,tidyOcr,imageType} from '../lib/doc-tools.mjs';
import {findCjkFont} from '../lib/fonts.mjs';
const tmp=async t=>{const d=await fs.mkdtemp(path.join(os.tmpdir(),'doctools-'));t.after(()=>fs.rm(d,{recursive:true,force:true}));return d;};
const PNG=Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4500000000049454e44ae426082','hex');
const JPG=Buffer.concat([Buffer.from([0xff,0xd8,0xff,0xe0]),Buffer.alloc(4000,7)]);
async function docx(dir,name,media){const z=new JSZip();z.file('[Content_Types].xml','<Types/>');z.file('word/document.xml','<w:document/>');for(const [n,b] of Object.entries(media))z.file('word/media/'+n,b);const p=path.join(dir,name);await fs.writeFile(p,await z.generateAsync({type:'nodebuffer'}));return p;}
test('识别结果整理：去掉中文、日文字之间的空格，英文单词间的空格保留；认得图片格式',()=>{
  assert.equal(tidyOcr('会 议 室 预 约\nお 問 い 合 わ せ\nHello World'),'会议室预约\nお問い合わせ\nHello World');
  assert.equal(imageType(PNG),'png');assert.equal(imageType(JPG),'jpg');assert.equal(imageType(Buffer.from('hello world')),'');
});
test('图片转文字：Windows 用系统 OCR，图片通过环境变量传给 PowerShell，缺语言时给出添加方法',async t=>{
  const calls=[];const exec=async(cmd,args,o)=>{calls.push({cmd,args,env:o.env});const files=o.env.OCR_FILES.split('|');return {stdout:JSON.stringify({lang:'ja-JP',langs:['ja-JP','en-US'],results:files.map(f=>({file:f,lines:['会 议 室','预约 10:00']}))})};};
  const tools=createDocTools({platform:'win32',exec,tmpRoot:await tmp(t)});
  const r=await tools.ocr({image:'data:image/png;base64,'+PNG.toString('base64'),lang:'zh'});
  assert.equal(r.text,'会议室\n预约 10:00');assert.equal(calls[0].cmd,'powershell.exe');assert.equal(calls[0].env.OCR_LANG,'zh-Hans-CN');assert.ok(!calls[0].args.join(' ').includes(calls[0].env.OCR_FILES));assert.match(r.note,/没有中文识别语言/);
  const none=createDocTools({platform:'win32',exec:async()=>({stdout:'{"error":"nolang","langs":[]}'}),tmpRoot:await tmp(t)});
  await assert.rejects(none.ocr({image:PNG.toString('base64')}),/光学字符识别/);
  await assert.rejects(tools.ocr({image:Buffer.from('not an image').toString('base64')}),/不是能识别的图片/);
  await assert.rejects(tools.ocr({}),/请先粘贴或选择/);
});
test('图片转文字：macOS 用 Vision，按从上到下、从左到右排好',async t=>{
  const exec=async(cmd,args)=>{assert.equal(cmd,'osascript');assert.equal(args[4],'ja-JP,en-US');return {stdout:JSON.stringify({results:[{file:args[5],items:[{t:'二行目',x:0.1,y:0.5,h:0.05},{t:'一行目',x:0.1,y:0.8,h:0.05},{t:'右',x:0.6,y:0.8,h:0.05}]}]})};};
  const r=await createDocTools({platform:'darwin',exec,tmpRoot:await tmp(t)}).ocr({image:PNG.toString('base64'),lang:'ja'});
  assert.equal(r.text,'一行目右\n二行目');
});
test('识别文档里的全部图片、提取图片到新文件夹（原文件不动）',async t=>{
  const d=await tmp(t),file=await docx(d,'报告.docx',{'image2.png':PNG,'image10.jpeg':JPG,'image1.png':PNG});
  const seen=[];const exec=async(cmd,args,o)=>{const files=o.env.OCR_FILES.split('|');seen.push(...files.map(f=>path.basename(f)));return {stdout:JSON.stringify({lang:'zh-Hans-CN',results:files.map((f,i)=>({file:f,lines:i===1?[]:['第'+(i+1)+'张']}))})};};
  const tools=createDocTools({platform:'win32',exec,tmpRoot:d});
  const r=await tools.ocr({path:file,lang:'zh'});assert.deepEqual(seen,['1-image1.png','2-image2.png','3-image10.jpeg']);assert.equal(r.text,'第1张\n\n第3张');assert.equal(r.count,3);
  const x=await tools.extractImages({path:file});assert.equal(x.count,3);assert.equal(x.dir,path.join(d,'报告-图片'));assert.deepEqual((await fs.readdir(x.dir)).sort(),['001-image1.png','002-image2.png','003-image10.jpeg']);
  assert.equal((await tools.extractImages({path:file})).dir,path.join(d,'报告-图片-2'));
  await assert.rejects(tools.extractImages({path:path.join(d,'a.wps')}),/不支持这种文件/);
  await assert.rejects(tools.extractImages({path:'相对/路径.docx'}),/保存到电脑上/);
});
test('压缩图片：只换掉变小的图片，另存为“-压缩”新文件，原文件不变',async t=>{
  const d=await tmp(t),file=await docx(d,'相册.docx',{'image1.jpeg':JPG,'image2.png':PNG}),orig=await fs.readFile(file);
  const exec=async(cmd,args,o)=>{assert.equal(o.env.SHRINK_Q,'60');for(const f of await fs.readdir(o.env.SHRINK_DIR))if(f.endsWith('.jpeg'))await fs.writeFile(path.join(o.env.SHRINK_DIR,f+'.out'),Buffer.concat([Buffer.from([0xff,0xd8,0xff,0xe0]),Buffer.alloc(500,1)]));return {stdout:'[]'};};
  const r=await createDocTools({platform:'win32',exec,tmpRoot:d}).compressImages({path:file,quality:60});
  assert.equal(r.changed,1);assert.equal(r.count,2);assert.ok(r.after<r.before);assert.equal(r.path,path.join(d,'相册-压缩.docx'));
  assert.deepEqual(await fs.readFile(file),orig);const z=await JSZip.loadAsync(await fs.readFile(r.path));assert.equal((await z.file('word/media/image1.jpeg').async('nodebuffer')).length,504);assert.deepEqual(await z.file('word/media/image2.png').async('nodebuffer'),PNG);
  const none=await createDocTools({platform:'win32',exec:async()=>({stdout:'[]'}),tmpRoot:d}).compressImages({path:file});assert.equal(none.changed,0);assert.equal(none.path,null);
});
test('PDF 转 Word：段落接回去、编号短行变标题，结果另存为同名 docx',{skip:!findCjkFont()},async t=>{
  const d=await tmp(t),f=findCjkFont(),pdfFile=path.join(d,'通知.pdf');
  await new Promise((resolve,reject)=>{const pdf=new PDFDocument({size:'A4',margin:50}),chunks=[];pdf.on('data',c=>chunks.push(c));pdf.on('end',async()=>{await fs.writeFile(pdfFile,Buffer.concat(chunks));resolve();});pdf.on('error',reject);pdf.font(f.file,f.family).fontSize(12);
    pdf.text('一、会议安排');pdf.text('本次会议定于下周三上午在三楼会议室召开，请各部门负责人准时参加，并提前准备好本季度的工作总结材料以及下一季度的计划。',{width:300});pdf.text('二、注意事项');pdf.text('请提前十分钟到场。');pdf.end();});
  const tools=createDocTools({platform:'linux',exec:async()=>{throw Error('不应调用');},tmpRoot:d});
  const r=await tools.pdfToWord({path:pdfFile,images:false});assert.equal(r.path,path.join(d,'通知.docx'));assert.equal(r.ocr,0);
  const xml=await (await JSZip.loadAsync(await fs.readFile(r.path))).file('word/document.xml').async('string');
  assert.match(xml,/Heading1[\s\S]*一、会议安排/);assert.match(xml,/本次会议定于下周三上午[^<]*材料以及下一季度的计划。/);assert.match(xml,/二、注意事项/);
  assert.equal((await tools.pdfToWord({path:pdfFile,images:false})).path,path.join(d,'通知-2.docx'));
  await assert.rejects(tools.pdfToWord({data:Buffer.from('hello').toString('base64'),name:'x.pdf'}),/不是有效的 PDF/);
});
test('段落重排：编号行、短行和句号结尾不接；英文之间补空格',()=>{
  const tools=createDocTools({platform:'linux'});
  const p=tools.paragraphs('第一章 总则\n这是一段很长很长的文字，第一行写到这里还没有\n结束，第二行接着写完。\n1. 第一项\n2. 第二项\nThe quick brown fox jumps over the lazy\ndog.');
  assert.deepEqual(p.map(x=>x.level),[1,0,2,2,0]);assert.equal(p[1].text,'这是一段很长很长的文字，第一行写到这里还没有结束，第二行接着写完。');assert.equal(p[4].text,'The quick brown fox jumps over the lazy dog.');
});
test('历史版本：列出备份文件夹里这个文件的副本（新的在前），自动备份只留最近 N 份',async t=>{
  const d=await tmp(t),b=path.join(d,'备份');await fs.mkdir(b);
  for(const n of ['合同-备份-20261101-090000.docx','合同-备份-20261115-101500-自动.docx','合同-备份-20261114-080000-自动.docx','合同-备份-20261113-080000-自动.docx','合同二-备份-20261101-090000.docx','合同-备份-20261101-090000.pdf'])await fs.writeFile(path.join(b,n),'x');
  const tools=createDocTools({platform:'linux'});const h=await tools.history({path:path.join(d,'合同.docx')},b);
  assert.deepEqual(h.items.map(x=>x.time),['2026-11-15 10:15:00','2026-11-14 08:00:00','2026-11-13 08:00:00','2026-11-01 09:00:00']);assert.equal(h.items[0].auto,true);assert.equal(h.items[3].auto,false);
  assert.equal(await tools.pruneAuto(path.join(d,'合同.docx'),b,2),1);assert.ok(!(await fs.readdir(b)).includes('合同-备份-20261113-080000-自动.docx'));assert.ok((await fs.readdir(b)).includes('合同-备份-20261101-090000.docx'));
});
test('服务接口：需要本机口令；自动备份同一文件 10 分钟内只留一次；未知功能 404',async t=>{
  const {createCoachServer}=await import('../server.mjs');const d=await tmp(t),b=path.join(d,'备份'),doc=path.join(d,'计划.docx');await fs.writeFile(doc,'v1');
  const server=createCoachServer({aiFile:path.join(d,'ai.json'),cardsFile:path.join(d,'c.json'),settingsFile:path.join(d,'s.json'),backupDir:b,privacyPluginDir:path.join(d,'plugins'),docToolsOptions:{platform:'linux',exec:async()=>({stdout:''})}});
  await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await (await fetch(base+'/api/session')).json()).token;
  const post=(p,body,k=token)=>fetch(base+p,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':k},body:JSON.stringify(body)});
  assert.equal((await post('/api/tools/history',{path:doc},'bad')).status,403);assert.equal((await post('/api/tools/nope',{})).status,404);
  const a=await (await post('/api/tools/auto-backup',{path:doc})).json();assert.match(a.backupPath,/计划-备份-\d{8}-\d{6}-自动\.docx$/);
  assert.equal((await (await post('/api/tools/auto-backup',{path:doc})).json()).skipped,'recent');assert.equal((await (await post('/api/tools/auto-backup',{path:'未保存'})).json()).skipped,'unsaved');
  const h=await (await post('/api/tools/history',{path:doc})).json();assert.equal(h.items.length,1);assert.equal(h.items[0].auto,true);
  const e=await post('/api/tools/ocr',{});assert.equal(e.status,422);assert.match((await e.json()).error,/请先粘贴或选择/);
});
test('PDF 转 Word：画了框线的表格还原成 Word 表格，表格的文字不再重复成段落',{skip:!findCjkFont()},async t=>{
  const d=await tmp(t),f=findCjkFont(),file=path.join(d,'销售.pdf');
  await new Promise((resolve,reject)=>{const pdf=new PDFDocument({size:'A4',margin:50}),chunks=[];pdf.on('data',c=>chunks.push(c));pdf.on('end',async()=>{await fs.writeFile(file,Buffer.concat(chunks));resolve();});pdf.on('error',reject);pdf.font(f.file,f.family).fontSize(11);
    pdf.text('一、销售情况',50,60);pdf.text('本季度各地区销售额如下表所示。',50,85);const x0=50,y0=110,cw=[120,100,100],rh=24,rows=[['地区','销售额','增长'],['华东','1,200','12%'],['华南','980','8%']];
    rows.forEach((r,i)=>{let x=x0;r.forEach((c,j)=>{pdf.text(c,x+6,y0+i*rh+6,{width:cw[j]-12});x+=cw[j];});});for(let i=0;i<=rows.length;i++)pdf.moveTo(x0,y0+i*rh).lineTo(x0+320,y0+i*rh).stroke();let x=x0;for(let j=0;j<=cw.length;j++){pdf.moveTo(x,y0).lineTo(x,y0+rows.length*rh).stroke();x+=cw[j]||0;}
    pdf.text('二、说明',50,200);pdf.text('以上数据未经审计。',50,225);pdf.end();});
  const r=await createDocTools({platform:'linux',tmpRoot:d}).pdfToWord({path:file,images:false});assert.equal(r.tables,1);
  const xml=await (await JSZip.loadAsync(await fs.readFile(r.path))).file('word/document.xml').async('string');
  assert.match(xml,/如下表所示。[\s\S]*<w:tbl>[\s\S]*地区[\s\S]*华南[\s\S]*980[\s\S]*<\/w:tbl>[\s\S]*二、说明/);assert.equal((xml.match(/华东/g)||[]).length,1);
});
test('表格位置：页面文字里对上表格的几行被替换，对不上的表格放在页末',()=>{
  const tools=createDocTools({platform:'linux'});
  const s=tools.withTables('前面\n地区 销售额\n华东 1,200\n后面',[[['地区','销售额'],['华东','1,200']],[['甲','乙'],['丙','丁']]]);
  assert.deepEqual(s.map(x=>x.table?'T':x.text),['前面','T','后面','T']);
});

test('PDF 转 Word：只有表格没有正文也能生成文件',async t=>{
 const d=await tmp(t),file=path.join(d,'table-only.pdf');
 await new Promise((resolve,reject)=>{
  const pdf=new PDFDocument({size:'A4',margin:50}),chunks=[];
  pdf.on('data',c=>chunks.push(c));pdf.on('error',reject);pdf.on('end',()=>fs.writeFile(file,Buffer.concat(chunks)).then(resolve,reject));
  [['Product','Quantity'],['Notebook','120'],['Pencil','300']].forEach((r,i)=>r.forEach((text,j)=>pdf.fontSize(11).text(text,56+j*150,56+i*30,{width:138})));
  for(let i=0;i<=3;i++)pdf.moveTo(50,50+i*30).lineTo(350,50+i*30).stroke();
  for(let j=0;j<=2;j++)pdf.moveTo(50+j*150,50).lineTo(50+j*150,140).stroke();pdf.end();
 });
 const r=await createDocTools({tmpRoot:d}).pdfToWord({path:file,images:false});assert.equal(r.tables,1);assert.equal(r.paragraphs,0);
 const xml=await (await JSZip.loadAsync(await fs.readFile(r.path))).file('word/document.xml').async('string');assert.match(xml,/<w:tbl>/);assert.match(xml,/Notebook/);
});
test('段落重排：封面、表单的“标签：内容”短行各自成段，写满一行的字段仍会接上下一行',()=>{
  const tools=createDocTools({platform:'linux'});
  const p=tools.paragraphs('FAX 送信票\n宛先：試験商事 ご担当者様\n差出人：サンプル株式会社\nFAX 番号：000-0000-0000\n件名：印刷と改行の確認\n送付資料をご確認ください。').map(x=>x.text);
  assert.deepEqual(p,['FAX 送信票','宛先：試験商事 ご担当者様','差出人：サンプル株式会社','FAX 番号：000-0000-0000','件名：印刷と改行の確認','送付資料をご確認ください。']);
  assert.equal(tools.paragraphs('地址：北京市测试区示例路十二号三单元某某大厦二十层二零零一室收件处请\n转交给张三先生本人签收，谢谢。').length,1);
});
