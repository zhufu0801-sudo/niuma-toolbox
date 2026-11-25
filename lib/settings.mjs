import {releaseUrl} from './release-info.mjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import {randomUUID} from 'node:crypto';
export function defaultSettings(backupDir=path.join(os.homedir(),'Documents','办公工具箱备份')){return {toolkit:{favorites:[],bindings:[]},hiddenIds:[],text:{font:'宋体',size:12,lineSpacing:1.5},table:{headerBackground:'#DDEBF7',headerColor:'#1F3864'},cards:{limit:40},backups:{directory:backupDir,auto:true},outputs:{directory:''},export:{wps:'docx',et:'xlsx',wpp:'pptx',drop:'docx'},launch:{autoOpenWps:false},updates:{url:''},ui:{language:'auto',beginner:false},pasteBar:{watch:false}};}
const copy=value=>structuredClone(value);
export function validateSettings(patch,current=defaultSettings()){
 if(!patch||typeof patch!=='object'||Array.isArray(patch)||Object.keys(patch).some(k=>!Object.hasOwn(current,k)))throw Error('设置格式不正确。');
 const next=copy(current);for(const [key,value] of Object.entries(patch)){if(key==='toolkit'){if(!value||typeof value!=='object'||Array.isArray(value)||Object.keys(value).some(k=>!['favorites','bindings'].includes(k)))throw Error('工具偏好格式不正确。');const t={...next.toolkit,...value};if(!Array.isArray(t.favorites)||t.favorites.length>100||t.favorites.some(x=>typeof x!=='string'||!/^[A-Za-z][A-Za-z0-9]{0,79}$/.test(x)))throw Error('收藏格式不正确。');if(!Array.isArray(t.bindings)||t.bindings.length>100)throw Error('快捷键数量不正确。');const ids=new Set(),keys=new Set(),tips=new Set();for(const b of t.bindings){if(!b||typeof b!=='object'||!['wps','et','wpp'].includes(b.host)||typeof b.id!=='string'||!/^[A-Za-z][A-Za-z0-9]{0,79}$/.test(b.id)||typeof b.key!=='string'||b.key&&!/^Ctrl\+Alt\+(Shift\+)?[A-Z0-9]$/.test(b.key)||typeof b.tip!=='string'||b.tip&&!/^[A-Z][A-Z0-9]?$/.test(b.tip))throw Error('快捷键格式不正确。');const id=b.host+':'+b.id,k=b.host+':'+b.key,t=b.host+':'+b.tip;if(ids.has(id)||b.key&&keys.has(k)||b.tip&&tips.has(t))throw Error('快捷键重复，请换一个组合。');ids.add(id);if(b.key)keys.add(k);if(b.tip)tips.add(t);}next.toolkit=t;continue;}if(key==='hiddenIds'){if(!Array.isArray(value)||value.length>100||value.some(id=>typeof id!=='string'||!/^[A-Za-z][A-Za-z0-9_-]{0,79}$/.test(id)))throw Error('功能开关格式不正确。');next.hiddenIds=[...new Set(value)].filter(id=>id!=='insertCard');continue;}if(!value||typeof value!=='object'||Array.isArray(value)||Object.keys(value).some(k=>!Object.hasOwn(current[key],k)))throw Error('设置项无法识别：'+key);Object.assign(next[key],value);}
 if(typeof next.text.font!=='string'||!next.text.font.trim()||next.text.font.length>80||/[\x00-\x1f]/.test(next.text.font))throw Error('请填写有效的正文字体。');next.text.font=next.text.font.trim();
 if(typeof next.text.size!=='number'||!Number.isFinite(next.text.size)||next.text.size<6||next.text.size>72)throw Error('正文字号应为 6～72。');
 if(typeof next.text.lineSpacing!=='number'||!Number.isFinite(next.text.lineSpacing)||next.text.lineSpacing<1||next.text.lineSpacing>3)throw Error('行距应为 1～3 倍。');
 for(const color of Object.values(next.table))if(typeof color!=='string'||!/^#[A-Fa-f0-9]{6}$/.test(color))throw Error('表头颜色格式不正确。');
 if(!Number.isInteger(next.cards.limit)||next.cards.limit<10||next.cards.limit>200)throw Error('卡片上限应为 10～200。');
 if(typeof next.backups.directory!=='string'||!path.isAbsolute(next.backups.directory)||next.backups.directory.length>300||/[\x00-\x1f]/.test(next.backups.directory))throw Error('请填写完整的备份文件夹路径。');
 if(typeof next.backups.auto!=='boolean')throw Error('自动备份选项格式不正确。');
 // 生成文件（挨拶状、寄件单、模板副本等）的保存位置；留空 = 文档\办公工具箱文件（日文界面为 オフィスツールボックス）。
 if(!next.outputs||typeof next.outputs.directory!=='string'||next.outputs.directory&&(!path.isAbsolute(next.outputs.directory)||next.outputs.directory.length>300||/[\x00-\x1f]/.test(next.outputs.directory)))throw Error('请填写完整的生成文件夹路径，或留空使用默认位置。');
 for(const target of Object.values(next.export))if(!['docx','xlsx','pptx','pdf'].includes(target))throw Error('默认保存格式无法识别。');
 if(typeof next.launch.autoOpenWps!=='boolean')throw Error('启动选项格式不正确。');
 next.updates.url=releaseUrl(next.updates.url);
 if(!['auto','zh','ja'].includes(next.ui.language))throw Error('界面语言无法识别。');
 if(typeof next.ui.beginner!=='boolean')throw Error('新手模式选项格式不正确。');
 // 粘贴栏自动记录复制内容：默认关闭。
 if(typeof next.pasteBar.watch!=='boolean')throw Error('粘贴栏选项格式不正确。');return next;
}
// 生成文件的默认文件夹：成品和备份分开放，客户一眼能找到。
export function outputDirectory(settings,{home=os.homedir(),locale=Intl.DateTimeFormat().resolvedOptions().locale}={}){
 const custom=settings&&settings.outputs&&settings.outputs.directory;if(custom)return custom;
 const lang=settings&&settings.ui&&settings.ui.language,ja=lang==='ja'||(lang!=='zh'&&/^ja/i.test(String(locale||'')));
 return path.join(home,'Documents',ja?'オフィスツールボックス':'办公工具箱文件');
}
export function createSettingsStore(file,{backupDir}={}){
 const defaults=defaultSettings(backupDir);let settings=copy(defaults),revision=0,initialized=false,loaded=false,chain=Promise.resolve();
 async function load(){if(loaded)return;try{if(file){const raw=JSON.parse(await fs.readFile(file,'utf8'));settings=validateSettings(raw.settings,defaults);revision=Number.isSafeInteger(raw.revision)?raw.revision:1;initialized=raw.initialized!==false;}}catch(e){if(e.code!=='ENOENT')throw Error('本机设置无法读取，已保留原文件，请备份后检查。');}loaded=true;}
 const snapshot=()=>({settings:copy(settings),defaults:copy(defaults),revision,initialized});
 function serial(fn){const result=chain.catch(()=>{}).then(fn);chain=result;return result;}
 async function commit(next){const at=Math.max(Date.now(),revision+1);if(file){await fs.mkdir(path.dirname(file),{recursive:true});const temp=file+'.'+randomUUID()+'.tmp';await fs.writeFile(temp,JSON.stringify({settings:next,revision:at,initialized:true},null,2),'utf8');await fs.rename(temp,file);}settings=next;revision=at;initialized=true;return snapshot();}
 return {
  get:()=>serial(async()=>{await load();return snapshot();}),
  update:(patch,expectedRevision)=>serial(async()=>{await load();if(expectedRevision!==undefined&&expectedRevision!==revision){const error=Error('设置已在另一个窗口修改，请刷新后再保存。');error.status=409;throw error;}return commit(validateSettings(patch,settings));}),
  migrate:legacy=>serial(async()=>{await load();if(initialized)return snapshot();return commit(validateSettings({hiddenIds:Array.isArray(legacy?.ids)?legacy.ids:[]},settings));})
 };
}
