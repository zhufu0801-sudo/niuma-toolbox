import {spawn,execFile} from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import {promisify} from 'node:util';
const execFileAsync=promisify(execFile);

// 按扩展名选 WPS 的程序：文字 wps.exe、表格 et.exe、演示 wpp.exe。
export function wpsProgramFor(file){
  const ext=path.extname(String(file)).toLowerCase();
  if(/^\.(docx?|dotx?|wps|wpt|rtf|txt)$/.test(ext))return 'wps.exe';
  if(/^\.(xlsx?|xlsm|xltx?|et|ett|csv)$/.test(ext))return 'et.exe';
  if(/^\.(pptx?|potx?|dps|dpt)$/.test(ext))return 'wpp.exe';
  return null;
}
async function wpsInstallRoot(){
  const keys=['HKCU\\Software\\Kingsoft\\Office\\6.0\\Common','HKLM\\Software\\Kingsoft\\Office\\6.0\\Common','HKLM\\Software\\WOW6432Node\\Kingsoft\\Office\\6.0\\Common'];
  for(const key of keys){try{const out=(await execFileAsync('reg.exe',['query',key,'/v','InstallRoot'],{windowsHide:true,timeout:5000})).stdout;const m=out.match(/InstallRoot\s+REG_(?:SZ|EXPAND_SZ)\s+(.+)/i);if(m)return m[1].trim();}catch{}}
  return null;
}
const launch=(cmd,args,msg)=>new Promise((resolve,reject)=>{const child=spawn(cmd,args,{detached:true,stdio:'ignore',windowsHide:false});child.once('error',()=>reject(Error(msg)));child.once('spawn',()=>{child.unref();resolve();});});

// 打开生成的文件。prefer='wps'（从 WPS 侧栏发起）时直接用 WPS 对应的程序打开，
// 不依赖 Windows 的文件关联——有些电脑 .xlsx 没设默认程序，会弹出“你要如何打开这个文件”。
// 找不到 WPS 或不是 WPS 的格式时，再按系统默认程序打开。
export async function openDocument(file,{prefer=null,platform=process.platform,findRoot=wpsInstallRoot,start=launch}={}){
  if(!path.isAbsolute(file))throw Error('文件路径必须为绝对路径。');
  const program=wpsProgramFor(file);
  if(prefer==='wps'&&program){
    if(platform==='win32'){const root=await findRoot().catch(()=>null);if(root){const exe=path.join(root,'office6',program);if(await fs.access(exe).then(()=>true,()=>false)){try{return await start(exe,[file],'WPS 启动失败。');}catch{}}}}
    else if(platform==='darwin'){try{return await start('open',['-b','com.kingsoft.wpsoffice.mac',file],'没有找到 Mac 版 WPS。');}catch{}}
  }
  const executable=platform==='win32'?'explorer.exe':platform==='darwin'?'open':null;
  if(!executable)throw Error('请手动打开生成的文件。');
  return start(executable,[file],'文件已保存，请手动打开。');
}
