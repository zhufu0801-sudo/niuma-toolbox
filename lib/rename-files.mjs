// 办公工具箱 · 批量重命名（2026-11-24）。先出预览，确认后再改；每次改名都记一份记录，可以一键改回。
// 规则：加前缀、加后缀、替换文字、按顺序编号。扩展名不动；新名字已被占用或不合法的跳过。
import fs from 'node:fs/promises';
import path from 'node:path';

const BAD=/[\\/:*?"<>|\x00-\x1f]/;
function stem(p){return path.basename(p,path.extname(p));}
export function newName(file,rule,index,count){const ext=path.extname(file),base=stem(file),r=rule||{};let name=base;
  if(r.mode==='prefix')name=String(r.text||'')+base;else if(r.mode==='suffix')name=base+String(r.text||'');
  else if(r.mode==='replace'){const find=String(r.find||'');if(!find)throw Error('请填写要替换的文字。');name=base.split(find).join(String(r.text||''));}
  else if(r.mode==='number'){const start=Number.isFinite(+r.start)&&String(r.start).trim()!==''?Math.trunc(+r.start):1,width=Math.max(1,Math.min(6,Number(r.width)||Math.max(2,String(start+Math.max(1,count||1)-1).length))),num=String(start+index).padStart(width,'0'),text=String(r.text||'');name=text.includes('#')?text.replace(/#+/,num):(text?text+num:num)+(r.keep?'-'+base:'');}
  else throw Error('请选择改名方式。');
  return name.trim()+ext;}
// 预览：返回每个文件的新名字和是否能改。
export async function planRename({paths,rule}){const list=[...new Set([].concat(paths||[]).map(String))];if(!list.length)throw Error('请先选择要改名的文件。');if(list.length>1000)throw Error('一次最多改 1000 个文件。');
  const taken=new Set(),plan=[];
  for(let i=0;i<list.length;i++){const from=list[i];if(!path.isAbsolute(from)){plan.push({from,to:'',skip:'位置无效'});continue;}const to=path.join(path.dirname(from),newName(from,rule,i,list.length));
    let skip='';const name=path.basename(to);if(!stem(to).trim())skip='新名字是空的';else if(BAD.test(name))skip='新名字里有 \\ / : * ? " < > | 这些不能用的字符';else if(name.length>200)skip='新名字太长';else if(to===from)skip='名字没有变化';else if(taken.has(to.toLowerCase()))skip='和前面的新名字重复';
    else if(!list.some(p=>p.toLowerCase()===to.toLowerCase())&&await fs.stat(to).then(()=>true,()=>false))skip='文件夹里已经有同名文件';
    if(!await fs.stat(from).then(s=>s.isFile(),()=>false))skip='找不到原文件';
    if(!skip)taken.add(to.toLowerCase());plan.push({from,to,skip});}
  return {plan,count:plan.filter(x=>!x.skip).length};}
// 执行：先全部改成临时名再改成新名，避免 A→B、B→A 这种互换时互相覆盖。记录写进 logDir。
export async function applyRename({paths,rule,logDir}){const {plan}=await planRename({paths,rule}),todo=plan.filter(x=>!x.skip),done=[],failed=[];
  const temp=todo.map((x,i)=>({...x,tmp:path.join(path.dirname(x.from),'.~rename-'+process.pid+'-'+Date.now()+'-'+i+path.extname(x.from))}));
  for(const x of temp){try{await fs.rename(x.from,x.tmp);x.moved=true;}catch(e){failed.push({from:x.from,error:e.code==='EBUSY'||e.code==='EPERM'?'文件正在被打开或没有权限':e.message});}}
  for(const x of temp.filter(x=>x.moved)){try{await fs.rename(x.tmp,x.to);done.push({from:x.from,to:x.to});}catch(e){await fs.rename(x.tmp,x.from).catch(()=>{});failed.push({from:x.from,error:e.message});}}
  let log='';if(done.length&&logDir){await fs.mkdir(logDir,{recursive:true});const d=new Date(),pad=n=>String(n).padStart(2,'0');log=path.join(logDir,'重命名记录-'+d.getFullYear()+pad(d.getMonth()+1)+pad(d.getDate())+'-'+pad(d.getHours())+pad(d.getMinutes())+pad(d.getSeconds())+'.json');await fs.writeFile(log,JSON.stringify({at:d.toISOString(),items:done},null,2),'utf8');}
  return {done,failed,skipped:plan.filter(x=>x.skip),log};}
// 改回：读最近一份记录，把新名字改回原来的名字（原名已被占用的跳过）。
export async function undoRename({logDir}){const files=(await fs.readdir(logDir).catch(()=>[])).filter(f=>/^重命名记录-\d{8}-\d{6}\.json$/.test(f)).sort();if(!files.length)throw Error('没有找到改名记录，无法改回。');
  const file=path.join(logDir,files[files.length-1]),rec=JSON.parse(await fs.readFile(file,'utf8')),back=[],failed=[];
  for(const x of [...rec.items].reverse()){if(await fs.stat(x.from).then(()=>true,()=>false)){failed.push({from:x.to,error:'原来的名字已经被别的文件占用'});continue;}try{await fs.rename(x.to,x.from);back.push(x);}catch(e){failed.push({from:x.to,error:e.code==='ENOENT'?'找不到文件（可能又被改过名或删掉了）':e.message});}}
  await fs.rename(file,file.replace(/\.json$/,'-已改回.json')).catch(()=>{});return {back,failed,log:file};}
