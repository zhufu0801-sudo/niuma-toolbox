import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {execFile,spawn} from 'node:child_process';
import {promisify} from 'node:util';
const exec=promisify(execFile);let cached=null,expires=0;
// ---- macOS：WPS 是 /Applications 里的 wpsoffice.app，加载项登记在应用沙盒容器里 ----
export const MAC_WPS_BUNDLE='com.kingsoft.wpsoffice.mac';
export function MAC_WPS_APPS(home=os.homedir()){return ['/Applications/wpsoffice.app','/Applications/WPS Office.app',path.join(home,'Applications','wpsoffice.app'),path.join(home,'Applications','WPS Office.app')];}
export function macRegistrationFiles(home=os.homedir()){return [path.join(home,'Library','Containers',MAC_WPS_BUNDLE,'Data','.kingsoft','wps','jsaddons','publish.xml'),path.join(home,'.kingsoft','wps','jsaddons','publish.xml')];}
// Office for Mac 的侧载文件夹：清单放进去，重开 Word / Excel / PowerPoint 后在“插入 → 加载项 → 我的加载项”里出现。
export function macOfficeSideloadDirs(home=os.homedir()){return ['Word','Excel','Powerpoint'].map(a=>({app:a.toLowerCase(),dir:path.join(home,'Library','Containers','com.microsoft.'+a,'Data','Documents','wef')}));}
export async function systemPaths(){
 if(cached&&Date.now()<expires)return cached;
 let data={};if(process.platform==='win32')try{
  const script="$ErrorActionPreference='Stop';[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false);$desktop=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders').GetValue('Desktop');$root=(Get-ItemProperty 'HKCU:\\Software\\Kingsoft\\Office\\6.0\\Common' -ErrorAction SilentlyContinue).InstallRoot;@{desktop=$desktop;wpsRoot=$root}|ConvertTo-Json -Compress";
  data=JSON.parse((await exec('powershell.exe',['-NoProfile','-NonInteractive','-Command',script],{windowsHide:true,timeout:7000,encoding:'utf8'})).stdout.trim());
 }catch{}
 const desktop=typeof data.desktop==='string'&&path.isAbsolute(data.desktop)?data.desktop:path.join(os.homedir(),'Desktop');
 const candidates=[];if(data.wpsRoot)candidates.push(path.join(data.wpsRoot,'office6','wps.exe'));const base=path.join(process.env.LOCALAPPDATA||'','Kingsoft','WPS Office');
 try{const dirs=(await fs.readdir(base,{withFileTypes:true})).filter(d=>d.isDirectory()&&/^\d/.test(d.name)).sort((a,b)=>b.name.localeCompare(a.name,undefined,{numeric:true}));for(const d of dirs)candidates.push(path.join(base,d.name,'office6','wps.exe'));}catch{}
 if(process.platform==='darwin')candidates.push(...MAC_WPS_APPS());
 let wpsProgram='';for(const file of candidates)try{const st=await fs.stat(file);if(st.isFile()||process.platform==='darwin'&&st.isDirectory()&&file.endsWith('.app')){wpsProgram=file;break;}}catch{}
 cached={desktop,wpsProgram};expires=Date.now()+30000;return cached;
}
export async function launchWps(program){if(!program)throw Error('没有找到 WPS，请先安装或修复 WPS。');await fs.access(program);const mac=process.platform==='darwin'&&program.endsWith('.app');return new Promise((resolve,reject)=>{const child=spawn(mac?'open':program,mac?[program]:[],{detached:true,stdio:'ignore',windowsHide:false,shell:false});child.on('error',()=>reject(Error('WPS 没有启动，请检查安装后重试。')));child.on('spawn',()=>{child.unref();resolve();});});}

// ---- WPS 版本识别（正版各版本适配）----
// 个人版通常装在“当前用户”目录（%LOCALAPPDATA%\Kingsoft\WPS Office\<版本>），专业版、企业版、政府版等常按电脑安装
// （Program Files\Kingsoft\WPS Office\<版本>），海外版和日本版（KINGSOFT）发行商和界面语言不同。
// 加载项登记在 %APPDATA%\kingsoft\wps\jsaddons\publish.xml；部分企业定制版需要在 office6\cfgs\oem.ini 的 [Support] 里开启 JsApiPlugin。
// 这里只读取这些信息并给出提示；需要管理员权限的修改不自动做。
export function classifyWps(info={}){
 const company=String(info.company||''),product=String(info.product||''),language=String(info.language||''),file=String(info.program||'');
 // 文件属性里的版本号常写成“12,1,0,29690”；安装目录名（…\\12.1.0.29690\\office6\\）更可靠，优先用它。
 const fromPath=(file.match(/\\(\d+(?:\.\d+){2,3})\\office6\\/i)||[])[1],fromInfo=(String(info.version||'').replace(/\s+/g,'').replace(/,/g,'.').match(/\d+(?:\.\d+){1,3}/)||[])[0];
 const version=fromPath||fromInfo||'';
 const major=Number(version.split('.')[0])||0;
 let region='unknown';
 // 实测日本版（WPS Office 2，11.2）的文件属性公司名也是珠海金山，只能靠注册表：默认语言 1041、更新服务器在 kingsoft.jp。
 const regLanguage=String(info.regLanguage||''),regUpdate=String(info.regUpdate||'');
 if(/^1041$/.test(regLanguage)||/kingsoft\.jp/i.test(regUpdate))region='jp';
 else if(/zhuhai|珠海|北京金山/i.test(company))region='cn';
 else if(/japan|ジャパン|キングソフト/i.test(company+product)||/日本語|japanese/i.test(language))region='jp';
 else if(/kingsoft/i.test(company))region='intl';
 const scope=/\\appdata\\local\\/i.test(file)?'user':file?'machine':'unknown';
 return {program:file,version,major,region,scope,product,company,language,regLanguage,active:info.active===true?true:info.active===false?false:null};
}
export function registrationStatus(xml,origin='http://127.0.0.1:39127'){
 const hosts={wps:false,et:false,wpp:false};if(typeof xml!=='string')return {found:false,hosts};
 const base=origin.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
 for(const h of Object.keys(hosts))hosts[h]=new RegExp(base+'/addon/'+h+'/','i').test(xml);
 return {found:true,hosts};
}
function decode(data){if(typeof data==='string')return data;if(!data||!data.length)return '';if(data[0]===0xff&&data[1]===0xfe)return data.subarray(2).toString('utf16le');if(data[0]===0xef&&data[1]===0xbb&&data[2]===0xbf)return data.subarray(3).toString('utf8');return data.toString('utf8');}
export function jsApiFlag(ini){ini=decode(ini);if(typeof ini!=='string')return null;const section=ini.match(/^\s*\[Support\]([\s\S]*?)(?=^\s*\[|$(?![\s\S]))/im);if(!section)return null;const line=section[1].match(/^\s*JsApiPlugin\s*=\s*(\S+)/im);return line?/^(true|1)$/i.test(line[1]):null;}
let envCache=null,envExpires=0;
export async function wpsEnvironment({run=exec,env=process.env,readFile=fs.readFile,readdir=fs.readdir,stat=fs.stat,platform=process.platform,cache=true}={}){
 if(cache&&envCache&&Date.now()<envExpires)return envCache;
 if(platform==='darwin'){const result=await macWpsEnvironment({readFile,stat,run});if(cache){envCache=result;envExpires=Date.now()+30000;}return result;}
 const candidates=new Set();const add=p=>{if(p)candidates.add(path.win32.normalize(p));};
 const scan=async base=>{try{const dirs=(await readdir(base,{withFileTypes:true})).filter(d=>d.isDirectory()&&/^\d/.test(d.name));for(const d of dirs)add(path.win32.join(base,d.name,'office6','wps.exe'));}catch{}};
 for(const base of [env.LOCALAPPDATA&&path.win32.join(env.LOCALAPPDATA,'Kingsoft','WPS Office'),env.ProgramFiles&&path.win32.join(env.ProgramFiles,'Kingsoft','WPS Office'),env['ProgramFiles(x86)']&&path.win32.join(env['ProgramFiles(x86)'],'Kingsoft','WPS Office')])if(base)await scan(base);
 let details=[];
 if(platform==='win32')try{
  const script="$ErrorActionPreference='SilentlyContinue';[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false);$roots=@();$meta=@{};foreach($k in 'HKCU:\\Software\\Kingsoft\\Office\\6.0\\Common','HKLM:\\SOFTWARE\\Kingsoft\\Office\\6.0\\Common','HKLM:\\SOFTWARE\\WOW6432Node\\Kingsoft\\Office\\6.0\\Common'){$g=Get-ItemProperty $k;$r=$g.InstallRoot;if($r){$roots+=[string]$r;$meta[([string]$r).TrimEnd('\\').ToLower()]=@{lang=[string]$g.DefaultLanguage;upd=[string]$g.UpdateSvr}}};$active=@($roots|ForEach-Object{$d=$_.TrimEnd('\\');if(Test-Path -LiteralPath (Join-Path $d 'wps.exe')){Join-Path $d 'wps.exe'}else{Join-Path $d 'office6\\wps.exe'}});$cands=@((ConvertFrom-Json $env:WPS_CANDIDATES)|ForEach-Object{[string]$_});$out=@();foreach($p in (@($active)+$cands|Select-Object -Unique)){if(Test-Path -LiteralPath $p){$v=(Get-Item -LiteralPath $p).VersionInfo;$out+=@{program=[string]$p;version=[string]$v.ProductVersion;product=[string]$v.ProductName;company=[string]$v.CompanyName;language=[string]$v.Language;active=[bool]($active -contains $p);regLanguage=$(foreach($m in $meta.Keys){if(([string]$p).ToLower().StartsWith($m)){$meta[$m].lang}});regUpdate=$(foreach($m in $meta.Keys){if(([string]$p).ToLower().StartsWith($m)){$meta[$m].upd}})}}};ConvertTo-Json -Compress -InputObject @($out)";
  const {stdout}=await run('powershell.exe',['-NoProfile','-NonInteractive','-Command',script],{windowsHide:true,timeout:9000,encoding:'utf8',env:{...env,WPS_CANDIDATES:JSON.stringify([...candidates])}});
  const parsed=JSON.parse(String(stdout).trim()||'[]');details=Array.isArray(parsed)?parsed:[parsed];
 }catch{}
 if(!details.length)for(const program of candidates)try{if((await stat(program)).isFile())details.push({program});}catch{}
 const seen=new Set(),installs=[];for(const d of details){const key=String(d.program||'').toLowerCase();if(!key||seen.has(key))continue;seen.add(key);const info=classifyWps(d);try{info.jsApiPlugin=jsApiFlag(await readFile(path.win32.join(path.win32.dirname(info.program),'cfgs','oem.ini')));}catch{info.jsApiPlugin=null;}installs.push(info);}
 // 自动更新后旧版本目录常会留下来：注册表指向的那个排最前，其余标成旧目录。
 if(installs.some(x=>x.active))for(const x of installs)if(!x.active)x.active=false;
 installs.sort((a,b)=>(b.active===true)-(a.active===true)||String(b.version).localeCompare(String(a.version),undefined,{numeric:true}));
 let registration={found:false,hosts:{wps:false,et:false,wpp:false},file:''};
 if(env.APPDATA){const file=path.win32.join(env.APPDATA,'kingsoft','wps','jsaddons','publish.xml');try{registration={...registrationStatus(decode(await readFile(file))),file};}catch{registration.file=file;}}
 const result={installs,registration};if(cache){envCache=result;envExpires=Date.now()+30000;}return result;
}

async function macWpsEnvironment({readFile=fs.readFile,stat=fs.stat,run=exec,home=os.homedir()}={}){
 const installs=[];for(const app of MAC_WPS_APPS(home)){try{if(!(await stat(app)).isDirectory())continue;}catch{continue;}
  let version='';try{const plist=String(await readFile(path.join(app,'Contents','Info.plist'),'utf8'));version=(plist.match(/<key>CFBundleShortVersionString<\/key>\s*<string>([^<]+)<\/string>/)||[])[1]||'';}catch{}
  installs.push({program:app,version,major:Number(version.split('.')[0])||0,region:/kingsoft\.jp/i.test(app)?'jp':'unknown',scope:app.startsWith(home)?'user':'machine',product:'WPS Office for Mac',company:'',language:'',regLanguage:'',active:installs.length?false:true,jsApiPlugin:null,platform:'mac'});}
 let registration={found:false,hosts:{wps:false,et:false,wpp:false},file:macRegistrationFiles(home)[0]};
 for(const file of macRegistrationFiles(home)){try{registration={...registrationStatus(decode(await readFile(file))),file};break;}catch{}}
 return {installs,registration};
}
// ---- 系统代理（排查用）----
// WPS 自己下载功能区图标等资源时走系统代理；代理不放行 127.0.0.1 时图标会显示成问号。这里只读取设置，不做修改。
export function proxyBypassesLocal(info={}){
 if(!info||!info.enabled)return true;
 const list=String(info.override||'').split(/[;\s]+/).filter(Boolean).map(s=>s.toLowerCase());
 return list.some(item=>item==='127.0.0.1'||item==='127.*'||item==='127.0.0.*'||item==='*');
}
let proxyCache=null,proxyExpires=0;
export async function proxySettings({run=exec,platform=process.platform,cache=true}={}){
 if(platform!=='win32')return {enabled:false,server:'',override:'',autoConfig:'',bypassLocal:true};
 if(cache&&proxyCache&&Date.now()<proxyExpires)return proxyCache;
 let raw={};try{
  const script="$ErrorActionPreference='SilentlyContinue';[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false);$g=Get-ItemProperty 'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings';@{enabled=[int]$g.ProxyEnable;server=[string]$g.ProxyServer;override=[string]$g.ProxyOverride;autoConfig=[string]$g.AutoConfigURL}|ConvertTo-Json -Compress";
  raw=JSON.parse(String((await run('powershell.exe',['-NoProfile','-NonInteractive','-Command',script],{windowsHide:true,timeout:7000,encoding:'utf8'})).stdout).trim()||'{}');
 }catch{}
 const info={enabled:Number(raw.enabled)===1,server:String(raw.server||'').slice(0,200),override:String(raw.override||'').slice(0,500),autoConfig:String(raw.autoConfig||'').slice(0,200)};
 info.bypassLocal=proxyBypassesLocal(info);
 proxyCache=info;proxyExpires=Date.now()+30000;return info;
}

// ---- Microsoft Office 检测（只读注册表）----
// 即点即用版看 ClickToRun\Configuration；MSI 版看 16.0/15.0 下各组件的 InstallRoot。Office 网页插件需要 2016 及以后（16.0）。
export function classifyOffice(raw={}){
 const ids=String(raw.productIds||'');const version=String(raw.version||'');const roots=Array.isArray(raw.msi)?raw.msi:[];
 let name='',major='';
 if(version){major=version.split('.')[0];
  if(/O365|M365|Microsoft365/i.test(ids))name='Microsoft 365';
  else if(/2024/.test(ids))name='Office 2024';else if(/2021/.test(ids))name='Office 2021';else if(/2019/.test(ids))name='Office 2019';else if(/2016/.test(ids))name='Office 2016';else name='Microsoft Office';}
 else if(roots.length){const best=roots.map(r=>String(r.ver||'')).sort().pop();major=best.split('.')[0];name=major==='16'?'Office 2016（MSI）':major==='15'?'Office 2013':'Microsoft Office';}
 if(!name)return {installed:false,supported:false,name:'',version:'',apps:[],language:''};
 const apps=[];const has=a=>new RegExp(a,'i').test(ids)||roots.some(r=>new RegExp(a,'i').test(r.app||''));
 if(raw.version){if(!/Visio|Project/i.test(ids)||/O365|ProPlus|Standard|HomeBusiness|HomeStudent|Excel|Word|PowerPoint|Professional/i.test(ids))apps.push('excel','word','powerpoint');}
 else{if(has('Excel'))apps.push('excel');if(has('Word'))apps.push('word');if(has('PowerPoint'))apps.push('powerpoint');}
 return {installed:true,supported:major==='16',name,version:version||(roots[0]&&roots[0].ver)||'',apps,language:String(raw.culture||''),platform:String(raw.platform||'')};
}
let officeCache=null,officeExpires=0;
export async function officeEnvironment({run=exec,platform=process.platform,cache=true}={}){
 if(platform==='darwin'){const apps=[],vers=[];for(const [id,name] of [['word','Microsoft Word'],['excel','Microsoft Excel'],['powerpoint','Microsoft PowerPoint']]){try{const plist=String(await fs.readFile('/Applications/'+name+'.app/Contents/Info.plist','utf8'));apps.push(id);vers.push((plist.match(/<key>CFBundleShortVersionString<\/key>\s*<string>([^<]+)<\/string>/)||[])[1]||'');}catch{}}
  if(!apps.length)return classifyOffice({});const v=vers.filter(Boolean).sort().pop()||'';return {installed:true,supported:!v||Number(v.split('.')[0])>=16,name:'Microsoft Office for Mac',version:v,apps,language:'',platform:'mac'};}
 if(platform!=='win32')return classifyOffice({});
 if(cache&&officeCache&&Date.now()<officeExpires)return officeCache;
 let raw={};try{
  const script="$ErrorActionPreference='SilentlyContinue';[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false);$c=Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration';$msi=@();foreach($b in @('HKLM:\\SOFTWARE\\Microsoft\\Office','HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Office')){foreach($v in @('16.0','15.0')){foreach($a in @('Excel','Word','PowerPoint')){$p=Get-ItemProperty ($b+'\\'+$v+'\\'+$a+'\\InstallRoot');if($p -and $p.Path){$msi+=@{ver=$v;app=$a}}}}};@{version=[string]$c.VersionToReport;productIds=[string]$c.ProductReleaseIds;culture=[string]$c.ClientCulture;platform=[string]$c.Platform;msi=$msi}|ConvertTo-Json -Compress -Depth 4";
  raw=JSON.parse(String((await run('powershell.exe',['-NoProfile','-NonInteractive','-Command',script],{windowsHide:true,timeout:8000,encoding:'utf8'})).stdout).trim()||'{}');
 }catch{}
 if(raw.msi&&!Array.isArray(raw.msi))raw.msi=[raw.msi];
 const info=classifyOffice(raw);officeCache=info;officeExpires=Date.now()+30000;return info;
}
