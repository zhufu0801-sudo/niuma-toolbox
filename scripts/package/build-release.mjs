// 打安装包：Windows 中文 / 日文、Mac 中文 / 日文（在项目根目录运行）。
//   node scripts/package/build-release.mjs --base <底包> --out <输出文件夹> [--only windows|mac]
// <底包> 提供运行组件：可以是上一版解压后的安装包文件夹（里面有 app\runtime\node.exe 和 app\node_modules），
//        也可以直接用已安装的目录（%LOCALAPPDATA%\Programs\WpsContentWorkbench，里面有 runtime 和 node_modules；不会读取 generated）。
// 程序代码（lib、public、plugins、server.mjs 等）一律取自当前项目；底包只用它的 runtime、node_modules 和 WPS-SDK-LICENSE.txt。
// 在 GitHub 的 macOS 机器上打 Mac 包时，底包就用项目自己（npm ci 之后），这样 node_modules 里是 Mac 用的组件；runtime 文件夹放一个空的即可。
// 日本邮编数据（lib/data/postal-jp.json.gz）超过 35 天会先自动从日本邮政官网更新。
// 日文版只把默认界面语言设成日文。Mac 版不带 Node.js（安装时从 nodejs.org 下载），并去掉只有 Windows 能用的图像组件。
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import JSZip from 'jszip';
import {prepareManuals} from './manuals.mjs';

const ROOT=process.cwd(),arg=k=>{const i=process.argv.indexOf('--'+k);return i>0?process.argv[i+1]:'';};
const pkg=JSON.parse(fs.readFileSync(path.join(ROOT,'package.json'),'utf8')),V=pkg.version;
let base=path.resolve(arg('base')||''),out=path.resolve(arg('out')||'release'),only=arg('only');
if(!arg('base'))throw Error('请用 --base 指定底包（上一版安装包文件夹或已安装目录）。');
if(fs.existsSync(path.join(base,'app','node_modules')))base=path.join(base,'app');
for(const need of ['node_modules','runtime'])if(!fs.existsSync(path.join(base,need)))throw Error('底包里找不到 '+need+'：'+base);
fs.mkdirSync(out,{recursive:true});
// Build fresh manuals from the same source revision as the application.
const {directory:manualDir}=prepareManuals({root:ROOT,out});
function copyManual(dst,ja){fs.copyFileSync(path.join(manualDir,ja?'manual-ja.pdf':'manual-zh.pdf'),path.join(dst,ja?'操作ガイド.pdf':'使用说明书.pdf'));}
function clearPackage(dst){const rel=path.relative(out,path.resolve(dst));if(!rel||rel==='..'||rel.startsWith('..'+path.sep)||path.isAbsolute(rel))throw Error('Unsafe package target: '+dst);fs.rmSync(dst,{recursive:true,force:true});}
// 日本邮编数据随安装包提供：没有或超过 35 天就先从日本邮政官网下载（2026-11-25）。下载失败时用现有的那份并提醒；一份都没有就停下。
{const {BUNDLED}=await import('../../lib/postal-jp.mjs');let age=Infinity;try{age=(Date.now()-fs.statSync(BUNDLED).mtimeMs)/864e5;}catch{}
 if(age>35){try{execFileSync(process.execPath,[path.join(ROOT,'scripts','data','update-postal-jp.mjs')],{stdio:'inherit'});}catch(e){if(!fs.existsSync(BUNDLED))throw Error('没有日本邮编数据，也没能下载。请先运行 node scripts/data/update-postal-jp.mjs（或加 --file 指定自己下载的 utf_ken_all.zip）。');console.warn('⚠ 没能更新日本邮编数据，安装包里用的是 '+Math.round(age)+' 天前的那份。');}}}
const cp=(a,b)=>fs.cpSync(a,b,{recursive:true,force:true});
function appDir(dst,{mac,ja}){
  fs.mkdirSync(dst,{recursive:true});
  for(const d of ['lib','public','plugins'])cp(path.join(ROOT,d),path.join(dst,d));
  cp(path.join(ROOT,'scripts','mcp'),path.join(dst,'scripts','mcp'));
  for(const f of ['server.mjs','package.json','package-lock.json'])cp(path.join(ROOT,f),path.join(dst,f));
  cp(path.join(ROOT,'LICENSE'),path.join(dst,'LICENSE'));
  for(const f of ['WPS-SDK-LICENSE.txt'])if(fs.existsSync(path.join(base,f)))cp(path.join(base,f),path.join(dst,f));
  cp(path.join(base,'node_modules'),path.join(dst,'node_modules'));
  // 底包是上一版做的，新加的依赖（如 0.16 的 kuromoji）底包里没有：从项目的 node_modules 补上（连同它自己的依赖）。
  for(const dep of Object.keys(pkg.dependencies||{})){const to=path.join(dst,'node_modules',dep);if(fs.existsSync(to))continue;const from=path.join(ROOT,'node_modules',dep);if(!fs.existsSync(from))throw Error('项目里还没装依赖 '+dep+'，请先运行 npm install。');cp(from,to);const meta=JSON.parse(fs.readFileSync(path.join(from,'package.json'),'utf8'));for(const sub of Object.keys(meta.dependencies||{})){const t2=path.join(dst,'node_modules',sub),f2=path.join(ROOT,'node_modules',sub);if(!fs.existsSync(t2)&&fs.existsSync(f2))cp(f2,t2);}}
  if(mac){const napi=path.join(dst,'node_modules','@napi-rs');if(fs.existsSync(napi))for(const d of fs.readdirSync(napi))if(/win32/.test(d))fs.rmSync(path.join(napi,d),{recursive:true,force:true});}
  else{cp(path.join(base,'runtime'),path.join(dst,'runtime'));cp(path.join(ROOT,'Launch.ps1'),path.join(dst,'Launch.ps1'));}
  if(ja){const p=path.join(dst,'lib','settings.mjs'),s=fs.readFileSync(p,'utf8');if(!s.includes("ui:{language:'auto'"))throw Error('lib/settings.mjs 里找不到默认语言设置，请检查。');fs.writeFileSync(p,s.replace("ui:{language:'auto'","ui:{language:'ja'"));}
  execFileSync(process.execPath,[path.join(ROOT,'scripts','package','slim.mjs'),dst],{stdio:'inherit'});
  execFileSync(process.execPath,[path.join(ROOT,'scripts','licenses','build-notices.mjs'),dst],{stdio:'inherit'});
}
async function zipDir(dir,file){
  const zip=new JSZip(),top=path.basename(dir);
  (function walk(d){for(const e of fs.readdirSync(d,{withFileTypes:true})){const p=path.join(d,e.name),rel=path.join(top,path.relative(dir,p)).split(path.sep).join('/');
    if(e.isDirectory())walk(p);else zip.file(rel,fs.readFileSync(p),{unixPermissions:/\.command$/.test(e.name)||/[\\/]runtime[\\/]bin[\\/]/.test(p)?0o755:0o644,date:new Date()});}})(dir);
  fs.writeFileSync(file,await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE',compressionOptions:{level:6},platform:'UNIX'}));
  console.log(path.basename(file),(fs.statSync(file).size/1048576).toFixed(1)+' MB');
}
// 更新说明：GitHub 上的仓库里可能没有这个文件（只在本机），没有就不放。
const P=path.join(ROOT,'scripts','package'),M=path.join(ROOT,'scripts','package-mac'),NOTES=path.join(ROOT,'办公工具箱-0.3.1更新说明.md');
for(const ja of [false,true]){const suf=ja?'-ja':'';
  if(only!=='mac'){const name=`OfficeToolbox-${V}-Windows-x64${suf}`,dst=path.join(out,name);clearPackage(dst);
    appDir(path.join(dst,'app'),{mac:false,ja});
    for(const f of ['Start.cmd','Uninstall.cmd','Initialize.ps1','Uninstall.ps1'])cp(path.join(P,f),path.join(dst,f));
    cp(path.join(P,'preset-example'),path.join(dst,'preset-example'));
    if(ja){cp(path.join(P,'README-ja.txt'),path.join(dst,'はじめにお読みください.txt'));cp(path.join(P,'docs','企業向け一括導入ガイド.md'),path.join(dst,'企業向け一括導入ガイド.md'));}
    else{cp(path.join(P,'README.txt'),path.join(dst,'README.txt'));cp(path.join(P,'docs','企业批量部署说明.md'),path.join(dst,'企业批量部署说明.md'));if(fs.existsSync(NOTES))cp(NOTES,path.join(dst,'更新说明.md'));}
    copyManual(dst,ja);
    await zipDir(dst,dst+'.zip');}
  if(only!=='windows'){const name=`OfficeToolbox-${V}-macOS${suf}`,dst=path.join(out,name);clearPackage(dst);
    appDir(path.join(dst,'app'),{mac:true,ja});
    for(const f of ['install.command','uninstall.command'])cp(path.join(M,f),path.join(dst,f));
    cp(path.join(M,ja?'README-ja.txt':'README.txt'),path.join(dst,ja?'はじめにお読みください.txt':'README.txt'));
    if(!ja&&fs.existsSync(NOTES))cp(NOTES,path.join(dst,'更新说明.md'));
    copyManual(dst,ja);
    await zipDir(dst,dst+'.zip');}
}
