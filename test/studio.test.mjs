import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import {normalizeStudio,planStudio,exportStudio,themes,wrapText} from '../lib/studio.mjs';
import {createCoachServer} from '../server.mjs';
const sample={title:'项目阶段汇报',subtitle:'阶段成果与计划',theme:'business',outline:'阶段成果\n第一条内容\n第二条内容\n\n下一步计划\n第三条内容'};
test('studio rejects incomplete or unbounded input instead of creating blank pages',()=>{
  for(const input of [{},{...sample,title:''},{...sample,theme:'__proto__'},{...sample,outline:'只有标题'},{...sample,outline:'标题\n'+ '字'.repeat(501)},{...sample,outline:Array(21).fill('章节\n内容').join('\n\n')}])assert.throws(()=>normalizeStudio(input));
});
test('automatic pagination retains all points and all layout objects fit inside the page',()=>{
  const points=Array.from({length:25},(_,i)=>'第'+i+'项：需要完整保留的内容');
  const plan=planStudio({...sample,outline:'任务列表\n'+points.join('\n')});
  assert.ok(plan.pages.length>3);
  const text=plan.pages.flatMap(p=>p.elements.filter(e=>e.type==='text').map(e=>e.text));
  for(const point of points)assert.ok(text.includes(point),point);
  for(const page of plan.pages)for(const item of page.elements){assert.ok(item.x>=0&&item.y>=0);assert.ok(item.x+item.w<=1280);assert.ok(item.y+item.h<=720);}
  const emoji='🙂中文 ABC';assert.equal(wrapText(emoji,50,20).join(''),emoji);
});
for(const theme of Object.keys(themes))test(`${theme}: exports a native editable PPTX with the same number of pages as the preview`,async()=>{
  const {buffer,plan}=await exportStudio({...sample,theme});const zip=await JSZip.loadAsync(buffer);
  const names=Object.keys(zip.files).filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n));assert.equal(names.length,plan.pages.length);
  const xml=(await Promise.all(names.map(n=>zip.file(n).async('string')))).join('');
  for(const text of ['项目阶段汇报','第一条内容','第三条内容'])assert.ok(xml.includes(text));
  assert.ok(xml.includes('<p:sp>'));assert.ok(!xml.includes('<p:pic>'));
});
test('studio HTTP endpoints authenticate, preview and export, then open only a newly created controlled file',async t=>{
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'office-studio-test-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));let opened;
  const server=createCoachServer({studioDir:dir,openPresentation:async filename=>{opened=filename;assert.ok((await fs.stat(filename)).size>1000);},fetchImpl:async()=>{throw Error('studio must not call external services');}});
  await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await(await fetch(base+'/api/session')).json()).token;
  const post=(action,data=sample,auth=token)=>fetch(base+'/api/studio/'+action,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':auth},body:JSON.stringify(data)});
  assert.equal((await post('open',sample,'wrong')).status,403);assert.equal(opened,undefined);
  assert.equal((await post('preview',{})).status,422);
  const preview=await(await post('preview')).json();assert.equal(preview.pages.length,4);assert.deepEqual(await fs.readdir(dir),[]);
  const download=await post('export');assert.equal(download.status,200);assert.ok((await download.arrayBuffer()).byteLength>1000);assert.deepEqual(await fs.readdir(dir),[]);
  const openedResult=await(await post('open',{...sample,filePath:'C:/should-not-overwrite.pptx'})).json();assert.equal(openedResult.opened,true);assert.equal(path.dirname(opened),dir);assert.equal(openedResult.filePath,opened);
});
test('a WPS launch failure keeps the new deck and returns its recovery location',async t=>{
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'office-studio-test-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));
  const server=createCoachServer({studioDir:dir,openPresentation:async()=>{throw Error('模拟未安装 WPS');}});await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await(await fetch(base+'/api/session')).json()).token;
  const response=await fetch(base+'/api/studio/open',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify(sample)});assert.equal(response.status,502);const result=await response.json();assert.ok(result.error.includes('文件已保存'));assert.ok((await fs.stat(result.filePath)).size>1000);
});
