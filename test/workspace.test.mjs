import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {importFile,exportFile,preflight,validateBlocks} from '../lib/workspace.mjs';
import {createCoachServer} from '../server.mjs';
const sample=[{type:'text',text:'项目计划：周五完成第一版。'},{type:'table',rows:[['姓名','任务'],['小林','界面设计'],['小陈','表格功能']]}];
for(const format of ['docx','pptx','xlsx','pdf']) test(`${format}: generates a readable file retaining Chinese content`,async()=>{
  const buffer=await exportFile(sample,format);
  assert.ok(buffer.length>1000);
  const result=await importFile('测试.'+format,buffer);
  const text=JSON.stringify(result.blocks);
  for(const value of ['项目计划','周五','小林','界面设计','小陈','表格功能']) assert.ok(text.includes(value),value);
  assert.ok(result.warnings.length);
  if(format==='docx') assert.deepEqual(result.blocks,sample);
});
test('unsupported, encrypted, oversized, empty and malformed files explain failure',async()=>{
  await assert.rejects(importFile('旧文件.doc',Buffer.from('fake')),/另存为/);
  await assert.rejects(importFile('加密.docx',Buffer.from('fake')),/加密或损坏/);
  await assert.rejects(importFile('大文件.pdf',Buffer.alloc(16*1024*1024)),/15 MB/);
  assert.throws(()=>validateBlocks([]),/没有可导出/);
  assert.throws(()=>validateBlocks([{type:'text',text:'  '}]),/内容为空/);
  assert.throws(()=>validateBlocks([{type:'table',rows:Array(201).fill(['a'])}]),/200/);
  assert.throws(()=>preflight(sample,'exe'),/请选择/);
});
test('blocks remain data, never spreadsheet formulas or macros',async()=>{
  const buffer=await exportFile([{type:'table',rows:[['=HYPERLINK("https://example.com")','00123','0']]}],'xlsx');
  const zip=await JSZip.loadAsync(buffer);
  assert.ok(!(await zip.file('xl/worksheets/sheet1.xml').async('string')).includes('<f>'));
  const result=await importFile('numbers.xlsx',buffer);
  assert.deepEqual(result.blocks[0].rows,[['=HYPERLINK("https://example.com")','00123','0']]);
});
test('numeric Word text including leading zeros survives import',async()=>{
  const buffer=await exportFile([{type:'text',text:'00123'},{type:'text',text:'0'}],'docx');
  assert.deepEqual((await importFile('numbers.docx',buffer)).blocks,[{type:'text',text:'00123\n0'}]);
});
test('Word merged columns retain alignment, content controls are read and paragraphs are grouped',async()=>{
  const zip=new JSZip();
  const p=text=>`<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  zip.file('word/document.xml',`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sdt><w:sdtContent>${p('模板正文')}${Array.from({length:110},(_,i)=>p('段落'+i)).join('')}</w:sdtContent></w:sdt><w:tbl><w:tr><w:tc>${p('第一列')}</w:tc><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${p('合并')}</w:tc><w:tc>${p('第四列')}</w:tc></w:tr><w:tr>${['甲','乙','丙','丁'].map(v=>'<w:tc>'+p(v)+'</w:tc>').join('')}</w:tr></w:tbl></w:body></w:document>`);
  const imported=await importFile('merged.docx',await zip.generateAsync({type:'nodebuffer'}));
  assert.equal(imported.blocks.length,2);assert.ok(imported.blocks[0].text.includes('模板正文'));assert.ok(imported.blocks[0].text.includes('段落109'));
  assert.deepEqual(imported.blocks[1].rows[0],['第一列','合并','','第四列']);
});
test('PPT table pagination repeats headers and retains every data row',async()=>{
  const rows=[['编号','任务'],...Array.from({length:24},(_,i)=>['编号'+i,'任务'+i])];
  const output=await exportFile([{type:'table',rows}],'pptx');
  const back=await importFile('paged.pptx',output);
  assert.equal(back.blocks.length,3);for(const b of back.blocks)assert.ok(b.text.startsWith('编号\n任务'));
  for(let i=0;i<24;i++)assert.ok(JSON.stringify(back.blocks).includes('任务'+i));
});
test('Excel and PDF preflight reject unrenderable dimensions before generation',()=>{
  assert.throws(()=>preflight([{type:'text',text:'长'.repeat(32768)}],'xlsx'),/32767/);
  assert.throws(()=>preflight([{type:'table',rows:[Array(11).fill('列')]}],'pdf'),/10 列/);
});
test('workspace API authenticates, checks, imports and downloads actual documents',async t=>{
  const server=createCoachServer();await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`;
  const token=(await (await fetch(base+'/api/session')).json()).token;
  const post=(action,body,auth=token)=>fetch(base+'/api/workspace/'+action,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':auth},body:JSON.stringify(body)});
  assert.equal((await post('export',{blocks:sample,target:'docx'},'wrong')).status,403);
  assert.equal((await post('check',{blocks:[],target:'pdf'})).status,422);
  assert.equal((await post('check',{blocks:sample,target:'pdf'})).status,200);
  const download=await post('export',{blocks:sample,target:'docx'});assert.equal(download.status,200);
  const imported=await post('import',{name:'sample.docx',data:Buffer.from(await download.arrayBuffer()).toString('base64')});
  assert.deepEqual((await imported.json()).blocks,sample);
  assert.equal((await fetch(base+'/workspace.html')).status,200);
});
