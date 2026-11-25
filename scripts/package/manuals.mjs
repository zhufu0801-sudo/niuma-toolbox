import fs from 'node:fs';
import path from 'node:path';
import {execFileSync} from 'node:child_process';
import {MISSING_FONT_EXIT} from '../docs/manual-fonts.mjs';
function validPdf(file){try{const b=fs.readFileSync(file);return b.subarray(0,5).toString()==='%PDF-'&&b.subarray(-1024).includes(Buffer.from('%%EOF'));}catch{return false;}}
export function prepareManuals({root,out,run=execFileSync,warn=console.warn}){
 fs.mkdirSync(out,{recursive:true});
 const stage=fs.mkdtempSync(path.join(out,'.manual-build-')),dest=path.join(out,'docs');
 let source=stage,fallback=false;
 try{
  try{run(process.execPath,[path.join(root,'scripts/docs/build-manual.mjs'),'--out',stage],{stdio:'inherit'});}
  catch(error){
   if(error.status!==MISSING_FONT_EXIT)throw error;
   source=[path.join(root,'release/docs'),path.join(root,'docs/manual/prebuilt')].find(dir=>['zh','ja'].every(lang=>validPdf(path.join(dir,`manual-${lang}.pdf`))));fallback=true;
   if(!source)throw Error('说明书缺少本机字体，且 release/docs 或 docs/manual/prebuilt 中没有可用的中文和日文 PDF。请先在有字体的电脑生成并随源码提供。');
   warn('⚠ 本机缺少说明书字体；本次使用 '+path.relative(root,source)+' 中已有的 PDF，内容可能不是最新版本。Mac 构建仍需在 macOS 验证。');
  }
  if(!['zh','ja'].every(lang=>validPdf(path.join(source,`manual-${lang}.pdf`))))throw Error('说明书生成结果缺少有效 PDF。');
  fs.mkdirSync(dest,{recursive:true});
  for(const lang of ['zh','ja','en'])for(const ext of ['pdf','docx']){
   const file=`manual-${lang}.${ext}`,from=path.join(source,file),to=path.join(dest,file);
   if(fs.existsSync(from)&&path.resolve(from)!==path.resolve(to))fs.copyFileSync(from,to);
  }
  if(!fallback&&fs.existsSync(path.join(stage,'build-report.json')))fs.copyFileSync(path.join(stage,'build-report.json'),path.join(dest,'build-report.json'));
  return {directory:dest,fallback};
 }finally{
  // stage is created by mkdtemp directly under the verified output directory.
  if(path.dirname(path.resolve(stage))!==path.resolve(out)||!path.basename(stage).startsWith('.manual-build-'))throw Error('Unsafe manual staging path');
  fs.rmSync(stage,{recursive:true,force:true});
 }
}
