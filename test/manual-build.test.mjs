import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fontCandidates,fontSpec,MISSING_FONT_EXIT} from '../scripts/docs/manual-fonts.mjs';
import {prepareManuals} from '../scripts/package/manuals.mjs';
import checks from '../scripts/docs/manual-checks.cjs';

test('manual steps keep quoted menu paths together and reject unbalanced lines',()=>{
 assert.deepEqual(checks.splitSteps('在“文件 → 数据脱敏”打开面板 → 核对'),['在“文件 → 数据脱敏”打开面板','核对']);
 assert.deepEqual(checks.splitSteps('「ファイル → マスキング」を開きます → 確認します'),['「ファイル → マスキング」を開きます','確認します']);
 for(const text of ['1. 在“文件','2. 数据脱敏”','1. 「ファイル','2. マスキング」','1. （コピー','2. 副本）'])assert.throws(()=>checks.validateManual(text),/unbalanced/);
 checks.validateManual('1. 在「文件」组点“数据脱敏”（先保存）\n| 表格例外（ |');
 for(const lang of ['zh','ja'])checks.validateManual(fs.readFileSync(new URL(`../docs/manual/${lang}.md`,import.meta.url),'utf8'),lang);
});

test('macOS manuals use installed Chinese font fallbacks and retain Japanese candidates',()=>{
 const options={platform:'darwin',env:{}},files=fontCandidates('zh',false,options).map(x=>x[0]);
 assert.ok(files.includes('/System/Library/Fonts/Hiragino Sans GB.ttc'));
 assert.ok(files.includes('/System/Library/Fonts/STHeiti Medium.ttc'));
 assert.ok(fontCandidates('ja',false,options).some(x=>x[0].includes('ヒラギノ')));
 for(const file of files.slice(1))assert.equal(fontSpec('zh',false,{...options,exists:f=>f===file,open:()=>({postscriptName:'test'})}).file,file);
 assert.throws(()=>fontSpec('zh',false,{...options,exists:()=>false}),{code:'MANUAL_FONT_MISSING'});
});

test('missing font uses preserved PDFs even if failed generation wrote partial files',()=>{
 const root=fs.mkdtempSync(path.join(os.tmpdir(),'manual-fallback-'));
 try{
  const cached=path.join(root,'release/docs');fs.mkdirSync(cached,{recursive:true});
  const pdf=Buffer.from('%PDF-1.4\nexisting fixture\n%%EOF\n');
  for(const lang of ['zh','ja'])fs.writeFileSync(path.join(cached,`manual-${lang}.pdf`),pdf);
  const warnings=[];const run=(_exe,args)=>{fs.writeFileSync(path.join(args[2],'manual-zh.pdf'),'partial');throw Object.assign(Error('missing font'),{status:MISSING_FONT_EXIT});};
  const result=prepareManuals({root,out:path.join(root,'release'),run,warn:x=>warnings.push(x)});
  assert.equal(result.fallback,true);assert.equal(warnings.length,1);
  assert.deepEqual(fs.readFileSync(path.join(cached,'manual-zh.pdf')),pdf);
  const other=prepareManuals({root,out:path.join(root,'other'),run,warn:()=>{}});
  assert.deepEqual(fs.readFileSync(path.join(other.directory,'manual-ja.pdf')),pdf);
  assert.ok(!fs.readdirSync(path.join(root,'release')).some(x=>x.startsWith('.manual-build-')));
  assert.throws(()=>prepareManuals({root,out:path.join(root,'bad'),run:()=>{throw Object.assign(Error('broken source'),{status:1});}}),/broken source/);
  fs.writeFileSync(path.join(cached,'manual-ja.pdf'),'partial');
  assert.throws(()=>prepareManuals({root,out:path.join(root,'missing'),run}),/没有可用/);
  const prebuilt=path.join(root,'docs/manual/prebuilt');fs.mkdirSync(prebuilt,{recursive:true});
  for(const lang of ['zh','ja'])fs.writeFileSync(path.join(prebuilt,`manual-${lang}.pdf`),pdf);
  const fromRepo=prepareManuals({root,out:path.join(root,'repo-only'),run,warn:()=>{}});
  assert.equal(fromRepo.fallback,true);assert.deepEqual(fs.readFileSync(path.join(fromRepo.directory,'manual-ja.pdf')),pdf);
 }finally{fs.rmSync(root,{recursive:true,force:true});}
});
