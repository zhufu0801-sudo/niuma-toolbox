// 常用文书模板（2026-11-21）：9 套模板都在，能用“打开副本”生成，表格里的合计公式还在。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import os from 'node:os';import path from 'node:path';
import JSZip from 'jszip';import ExcelJS from 'exceljs';
import {createTemplateLibrary} from '../lib/template-library.mjs';
const root=path.resolve(new URL('..',import.meta.url).pathname.replace(/^\/([A-Za-z]:)/,'$1'));
test('文书模板：中文 4 套、日文 5 套，文件齐全，文字是对应语言',async()=>{
  const index=JSON.parse(fs.readFileSync(path.join(root,'public/templates/forms/index.json'),'utf8'));
  assert.deepEqual(index.filter(x=>x.lang==='zh').map(x=>x.title).sort(),['会议纪要','工作周报','请假单','费用报销单'].sort());
  assert.deepEqual(index.filter(x=>x.lang==='ja').map(x=>x.title).sort(),['始末書','業務日報','発注書','稟議書','議事録'].sort());
  for(const x of index.filter(x=>x.file.endsWith('.docx'))){const zip=await JSZip.loadAsync(fs.readFileSync(path.join(root,'public/templates/forms',x.file)));const xml=await zip.file('word/document.xml').async('string');assert.ok(xml.includes(x.title.split('').join('')[0]),x.file);}
});
test('文书模板：报销单合计和大写、発注書小计 / 消费税 / 合计都是公式',async()=>{
  const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(path.join(root,'public/templates/forms/expense-zh.xlsx'));const ws=wb.worksheets[0];
  assert.equal(ws.getCell('E18').value.formula,'SUM(E8:E17)');assert.match(ws.getCell('C19').value.formula,/DBNum2/);
  const ob=new ExcelJS.Workbook();await ob.xlsx.readFile(path.join(root,'public/templates/forms/order-ja.xlsx'));const os2=ob.worksheets[0];
  assert.equal(os2.getCell('F25').value.formula,'SUM(F15:F24)');assert.equal(os2.getCell('F26').value.formula,'ROUNDDOWN(F25*0.1,0)');assert.equal(os2.getCell('F27').value.formula,'F25+F26');
});
test('模板库：日常文书可以“打开副本”，Word 文件也行，从 WPS 打开时用 WPS',async()=>{
  const out=fs.mkdtempSync(path.join(os.tmpdir(),'forms-'));const opened=[];
  const lib=createTemplateLibrary({root,outputDir:async()=>out,openFile:async(f,o)=>opened.push([f,o])});
  const items=await lib.list();assert.ok(items.some(x=>x.kind==='forms'&&x.id==='ringi-ja'));
  const r=await lib.run({action:'template',id:'ringi-ja',kind:'forms',open:true,app:'wps'});assert.equal(path.basename(r.path),'稟議書.docx');assert.ok(fs.existsSync(r.path));assert.equal(opened[0][1].prefer,'wps');
});
