// 办公工具箱 · Mac 版安装 / 卸载（由“安装.command”“卸载.command”调用）。
// 做的事：登记到 Mac 版 WPS（publish.xml）、给 Microsoft Office for Mac 侧载插件清单（需要本机 https 证书）、设置登录后自动启动本机服务。
// 原则：已有的 publish.xml 先备份再改，只增删办公工具箱自己的三条；证书只给本机 localhost 用；不改系统设置。
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { execFile } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { MAC_WPS_BUNDLE, MAC_WPS_APPS, macRegistrationFiles, macOfficeSideloadDirs } from './windows.mjs';

export const LABEL='com.officetoolbox.server';
const ENTRIES=[['WpsCoachWriter','wps'],['WpsCoachSheets','et'],['WpsCoachSlides','wpp']];
const ja=/^ja/i.test(process.env.LANG||process.env.LC_ALL||'')||process.env.TOOLBOX_LANG==='ja';
const say=(zh,jp)=>console.log(ja?jp:zh);
const run=(cmd,args,opts={})=>new Promise((resolve,reject)=>execFile(cmd,args,{encoding:'utf8',timeout:opts.timeout||120000},(e,out,err)=>e?reject(Object.assign(e,{stderr:err})):resolve(out)));
const exists=p=>fs.stat(p).then(()=>true,()=>false);
const ours=new RegExp('[ \\t]*<jspluginonline\\b[^>]*\\bname="('+ENTRIES.map(e=>e[0]).join('|')+')"[^>]*?(\\/>|>[\\s\\S]*?<\\/jspluginonline>)[ \\t]*(\\r?\\n)?','g');
// 在已有的 publish.xml 里换上办公工具箱的三条（别的加载项原样保留）。
export function mergePublishXml(xml,base='http://127.0.0.1:39127'){
  let s=String(xml||'').replace(/^﻿/,'');
  if(!s.trim())s='<?xml version="1.0" encoding="UTF-8"?>\n<jsplugins>\n</jsplugins>\n';
  if(!/<jsplugins\b/.test(s))throw Error('WPS 的 publish.xml 格式和预期不同，原文件没有改动。');
  s=s.replace(/<jsplugins\s*\/>/,'<jsplugins>\n</jsplugins>').replace(ours,'');
  const items=ENTRIES.map(([name,type])=>`  <jspluginonline name="${name}" type="${type}" install="${base}" url="${base}/addon/${type}/" enable="enable_dev"/>\n`).join('');
  return s.replace(/<\/jsplugins>/,items+'</jsplugins>');
}
export function removePublishEntries(xml){return String(xml||'').replace(ours,'');}
export function officeManifest(template,{id,version,base='https://localhost:39128'}){return template.replaceAll('{{ID}}',id).replaceAll('{{BASE}}',base).replaceAll('{{VERSION}}',version).replace('{{REQUIREMENTS}}','').replace(/\{\{RUNTIMES\}\}/g,'').replace('{{EXTENDED}}','');}
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
export function launchAgent({node,appDir}){
  const log=path.posix.join(appDir,'generated','server.log');
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>${LABEL}</string>
  <key>ProgramArguments</key><array><string>${esc(node)}</string><string>${esc(path.posix.join(appDir,'server.mjs'))}</string></array>
  <key>WorkingDirectory</key><string>${esc(appDir)}</string>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><dict><key>SuccessfulExit</key><false/></dict>
  <key>ThrottleInterval</key><integer>10</integer>
  <key>StandardOutPath</key><string>${esc(log)}</string>
  <key>StandardErrorPath</key><string>${esc(log)}</string>
</dict></plist>
`;
}
export const CERT_CONFIG=`[req]
distinguished_name=dn
x509_extensions=v3
prompt=no
[dn]
CN=OfficeToolbox localhost
[v3]
subjectAltName=DNS:localhost,IP:127.0.0.1
basicConstraints=critical,CA:false
keyUsage=critical,digitalSignature,keyEncipherment
extendedKeyUsage=serverAuth
`;

async function writeAtomic(file,text){await fs.mkdir(path.dirname(file),{recursive:true});const tmp=file+'.'+randomUUID()+'.tmp';await fs.writeFile(tmp,text,'utf8');await fs.rename(tmp,file);}
export async function detect(home=os.homedir()){
  let wps=null;for(const app of MAC_WPS_APPS(home))if(await exists(app)){wps=app;break;}
  const office=[];for(const [id,name] of [['word','Microsoft Word'],['excel','Microsoft Excel'],['powerpoint','Microsoft PowerPoint']])if(await exists('/Applications/'+name+'.app'))office.push(id);
  return {wps,office};
}
export async function registerWps(home=os.homedir()){
  const [container,plain]=macRegistrationFiles(home);
  const targets=[plain];if(await exists(path.join(home,'Library','Containers',MAC_WPS_BUNDLE)))targets.unshift(container);
  for(const file of targets){let old='';try{old=await fs.readFile(file,'utf8');await fs.copyFile(file,file+'.toolbox-backup-'+Date.now());}catch(e){if(e.code!=='ENOENT')throw e;}await writeAtomic(file,mergePublishXml(old));}
  say('已登记到 WPS Office（'+targets.length+' 处）。','WPS Office に登録しました（'+targets.length+' か所）。');return targets;
}
export async function registerOffice(appDir,home=os.homedir(),exec=run){
  const gen=path.join(appDir,'generated');await fs.mkdir(gen,{recursive:true});
  const cert=path.join(gen,'office-cert.pem'),key=path.join(gen,'office-key.pem');
  if(!(await exists(cert))||!(await exists(key))){
    const cfg=path.join(gen,'office-cert.cnf');await fs.writeFile(cfg,CERT_CONFIG);
    await exec('openssl',['req','-x509','-newkey','rsa:2048','-nodes','-sha256','-days','1825','-keyout',key,'-out',cert,'-config',cfg]);await fs.rm(cfg,{force:true});await fs.chmod(key,0o600);
    say('正在信任本机证书：macOS 会要求输入这台 Mac 的登录密码（只用于 localhost）。','このパソコン専用の証明書を信頼します。macOS のログインパスワードを求められたら入力してください（localhost 専用）。');
    try{await exec('security',['add-trusted-cert','-r','trustRoot','-p','ssl','-k',path.join(home,'Library','Keychains','login.keychain-db'),cert],{timeout:300000});}
    catch(e){say('没能自动信任证书（可能取消了密码输入）。Office 侧栏打不开时，请双击 '+cert+'，在“钥匙串访问”里把它设为“始终信任”。','証明書を自動で信頼できませんでした。Office のサイドバーが開かないときは '+cert+' をダブルクリックし、「キーチェーンアクセス」で「常に信頼」にしてください。');}
  }
  const idFile=path.join(gen,'office-id.txt');let id='';try{id=(await fs.readFile(idFile,'utf8')).trim();}catch{}if(!id){id=randomUUID();await fs.writeFile(idFile,id);}
  const version=JSON.parse(await fs.readFile(path.join(appDir,'package.json'),'utf8')).version+'.0';
  const manifest=officeManifest(await fs.readFile(path.join(appDir,'public','office','manifest-template.xml'),'utf8'),{id,version});
  await writeAtomic(path.join(gen,'office-manifest.xml'),manifest);
  const written=[];for(const {dir} of macOfficeSideloadDirs(home)){if(!(await exists(path.dirname(path.dirname(dir)))))continue;const f=path.join(dir,'OfficeToolbox-'+id+'.xml');await writeAtomic(f,manifest);written.push(f);}
  if(written.length)say('已放入 Microsoft Office 的加载项文件夹（'+written.length+' 个程序）。请重新打开 Word / Excel / PowerPoint，在“插入 → 加载项 → 我的加载项”里打开“办公工具箱”。','Microsoft Office のアドインフォルダーに配置しました（'+written.length+' 個のアプリ）。Word / Excel / PowerPoint を開き直し、「挿入 → アドイン → 個人用アドイン」から開いてください。');
  else say('没有找到 Word / Excel / PowerPoint 的数据文件夹：请先打开一次 Office，再重新运行安装。','Word / Excel / PowerPoint のデータフォルダーが見つかりません。一度 Office を起動してから、もう一度インストールしてください。');
  return written;
}
export async function installAgent(appDir,node,home=os.homedir()){
  const plist=path.join(home,'Library','LaunchAgents',LABEL+'.plist');await writeAtomic(plist,launchAgent({node,appDir}));
  const uid=String(process.getuid?process.getuid():'');await run('launchctl',['bootout','gui/'+uid,plist]).catch(()=>{});
  try{await run('launchctl',['bootstrap','gui/'+uid,plist]);}catch{await run('launchctl',['load','-w',plist]).catch(()=>{});}
  say('已设置为登录后自动启动本机服务。','ログイン後にローカルサービスが自動で起動するように設定しました。');return plist;
}
export async function uninstall(appDir,home=os.homedir(),exec=run){
  const plist=path.join(home,'Library','LaunchAgents',LABEL+'.plist'),uid=String(process.getuid?process.getuid():'');
  await exec('launchctl',['bootout','gui/'+uid,plist]).catch(()=>exec('launchctl',['unload',plist]).catch(()=>{}));await fs.rm(plist,{force:true});
  for(const file of macRegistrationFiles(home)){let old;try{old=await fs.readFile(file,'utf8');}catch{continue;}const next=removePublishEntries(old);if(next!==old){await fs.copyFile(file,file+'.toolbox-backup-'+Date.now());await writeAtomic(file,next);}}
  for(const {dir} of macOfficeSideloadDirs(home)){let names=[];try{names=await fs.readdir(dir);}catch{continue;}for(const n of names)if(/^OfficeToolbox-.*\.xml$/.test(n))await fs.rm(path.join(dir,n),{force:true});}
  const cert=path.join(appDir,'generated','office-cert.pem');if(await exists(cert)){try{const sha=(await exec('openssl',['x509','-noout','-fingerprint','-sha1','-in',cert])).split('=').pop().replace(/:/g,'').trim();await exec('security',['delete-certificate','-Z',sha,path.join(home,'Library','Keychains','login.keychain-db')]);}catch{}}
  say('已从 WPS 和 Microsoft Office 中移除办公工具箱，并停止了本机服务。你的设置和备份文件没有删除。','WPS と Microsoft Office からツールボックスを取り除き、ローカルサービスを停止しました。設定とバックアップは削除していません。');
}

if(process.argv[1]&&process.argv[1].endsWith('mac-install.mjs')){
  const [cmd,...rest]=process.argv.slice(2);const opt=k=>{const i=rest.indexOf('--'+k);return i>=0?rest[i+1]:'';};
  const appDir=path.resolve(opt('app')||path.join(path.dirname(new URL(import.meta.url).pathname),'..'));
  (async()=>{
    if(cmd==='detect'){console.log(JSON.stringify(await detect()));return;}
    if(cmd==='uninstall'){await uninstall(appDir);return;}
    if(cmd!=='install')throw Error('用法：node lib/mac-install.mjs install|uninstall|detect --app <文件夹> [--targets wps,office] [--node <node 路径>]');
    const targets=(opt('targets')||'wps,office').split(',').map(s=>s.trim()).filter(Boolean),node=opt('node')||process.execPath;
    await fs.mkdir(path.join(appDir,'generated'),{recursive:true});
    if(targets.includes('wps'))await registerWps();
    if(targets.includes('office'))await registerOffice(appDir);
    await installAgent(appDir,node);
    await fs.writeFile(path.join(appDir,'generated','targets.json'),JSON.stringify({targets,platform:'darwin'}));
  })().catch(e=>{console.error((ja?'エラー：':'出错了：')+e.message);process.exit(1);});
}
