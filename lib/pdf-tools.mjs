// PDF 工具（2026-11-21）：合并、拆分、提取 / 删除页面、PDF 转图片、图片转 PDF。
// 原文件一律不改，结果放进“生成文件”文件夹（重名自动加编号）。PDF 的读写用 pdf-lib（MIT，lib/vendor 里自带），
// 转图片用已有的 pdf-parse（pdf.js + canvas）。
import fs from 'node:fs/promises';import path from 'node:path';import {createRequire} from 'node:module';
import {ensurePdfEnv} from './pdf-env.mjs';
const require=createRequire(import.meta.url);
let lib=null;function pdfLib(){return lib||(lib=require('./vendor/pdf-lib.min.cjs'));}

export const LIMITS={files:50,fileMB:200,totalMB:500,pages:2000,images:200,imageMB:40,renderPages:100};
const safe=s=>String(s||'').replace(/[\\/:*?"<>|\x00-\x1f]+/g,' ').trim().slice(0,60)||'PDF';
async function freeName(dir,base,ext){let name=base+ext,i=2;while(await fs.stat(path.join(dir,name)).then(()=>true,()=>false))name=base+'-'+(i++)+ext;return path.join(dir,name);}
async function readInput(file,{exts=['.pdf'],limitMB=LIMITS.fileMB}={}){const f=String(file||'').trim();if(!f||!path.isAbsolute(f))throw Error('请选择已保存的文件（要完整位置，例如 C:\\资料\\合同.pdf）。');
  if(!exts.includes(path.extname(f).toLowerCase()))throw Error(exts.includes('.pdf')?'“'+path.basename(f)+'”不是 PDF 文件。':'“'+path.basename(f)+'”不是 JPG 或 PNG 图片。');
  let st;try{st=await fs.stat(f);}catch{throw Error('找不到文件：'+f);}if(!st.isFile())throw Error('找不到文件：'+f);if(st.size>limitMB*1024*1024)throw Error('“'+path.basename(f)+'”超过 '+limitMB+' MB，请先压缩或拆小。');return {file:f,buffer:await fs.readFile(f)};}
async function loadPdf(file,buffer){const {PDFDocument}=pdfLib();try{return await PDFDocument.load(buffer,{updateMetadata:false});}catch(e){
  if(/encrypt/i.test(String(e&&e.message)))throw Error('“'+path.basename(file)+'”设置了密码或禁止编辑，请先在 PDF 阅读器里解除后再试。');throw Error('“'+path.basename(file)+'”打不开，可能不是完整的 PDF 文件。');}}
async function savePdf(doc){return Buffer.from(await doc.save({useObjectStreams:true}));}

// 页码：1-3,5,8- ；也认全角数字、中文逗号、顿号、“～”。8- 表示第 8 页到最后。
export function parsePages(text,total){const s=String(text||'').replace(/[０-９]/g,c=>String.fromCharCode(c.charCodeAt(0)-0xFEE0)).replace(/[，、;；\s]+/g,',').replace(/[～~—–－]/g,'-').replace(/^,|,$/g,'');
  if(!s)throw Error('请填写页码，例如：1-3,5,8-');const out=[],seen=new Set();
  for(const part of s.split(',')){if(!part)continue;const m=part.match(/^(\d*)-(\d*)$|^(\d+)$/);if(!m||(m[1]===''&&m[2]==='')&&!m[3])throw Error('页码“'+part+'”看不懂，请写成 1-3,5,8- 这样。');
    let a,b;if(m[3]){a=b=Number(m[3]);}else{a=m[1]?Number(m[1]):1;b=m[2]?Number(m[2]):total;}
    if(a<1||b<1)throw Error('页码要从 1 开始。');if(a>total)throw Error('这个 PDF 只有 '+total+' 页，没有第 '+a+' 页。');if(b>total)b=total;if(a>b)[a,b]=[b,a];
    for(let p=a;p<=b;p++)if(!seen.has(p)){seen.add(p);out.push(p);}}
  if(!out.length)throw Error('请填写页码，例如：1-3,5,8-');return out;}
const rangeLabel=pages=>{const r=[];for(let i=0;i<pages.length;i++){let j=i;while(j+1<pages.length&&pages[j+1]===pages[j]+1)j++;r.push(j>i?pages[i]+'-'+pages[j]:String(pages[i]));i=j;}return r.join(',');};

// 合并：按给出的顺序首尾相接。
export async function mergePdfs({paths,outputDir}){const list=[].concat(paths||[]).map(p=>String(p).trim()).filter(Boolean);if(list.length<2)throw Error('请至少选择 2 个 PDF 文件。');if(list.length>LIMITS.files)throw Error('一次最多合并 '+LIMITS.files+' 个文件。');
  const {PDFDocument}=pdfLib(),out=await PDFDocument.create();let total=0,size=0;
  for(const p of list){const {file,buffer}=await readInput(p);size+=buffer.length;if(size>LIMITS.totalMB*1024*1024)throw Error('文件加起来超过 '+LIMITS.totalMB+' MB，请分几次合并。');
    const src=await loadPdf(file,buffer),pages=await out.copyPages(src,src.getPageIndices());pages.forEach(pg=>out.addPage(pg));total+=pages.length;if(total>LIMITS.pages)throw Error('合并后超过 '+LIMITS.pages+' 页，请分几次合并。');}
  await fs.mkdir(outputDir,{recursive:true});const target=await freeName(outputDir,safe(path.basename(list[0],path.extname(list[0])))+'-合并（'+list.length+' 个文件）','.pdf');
  await fs.writeFile(target,await savePdf(out));return {path:target,files:list.length,pages:total};}

// 提取（keep）或删除（remove）指定的页，存成新文件。
export async function pickPages({path:p,pages,mode='keep',outputDir}){const {file,buffer}=await readInput(p),src=await loadPdf(file,buffer),total=src.getPageCount(),chosen=parsePages(pages,total);
  const keep=mode==='remove'?src.getPageIndices().map(i=>i+1).filter(n=>!chosen.includes(n)):chosen;if(!keep.length)throw Error('删掉这些页以后就没有页面了。');
  const {PDFDocument}=pdfLib(),out=await PDFDocument.create();(await out.copyPages(src,keep.map(n=>n-1))).forEach(pg=>out.addPage(pg));
  await fs.mkdir(outputDir,{recursive:true});const base=safe(path.basename(file,path.extname(file)));
  const target=await freeName(outputDir,base+(mode==='remove'?'-删除第'+rangeLabel(chosen.slice().sort((a,b)=>a-b))+'页':'-第'+rangeLabel(chosen)+'页'),'.pdf');
  await fs.writeFile(target,await savePdf(out));return {path:target,pages:keep.length,total};}

// 拆分：every = 每 N 页一个文件（1 = 每页一个）；或者 ranges = “1-3,4-10,11-” 按段拆。
export async function splitPdf({path:p,every=1,ranges='',outputDir}){const {file,buffer}=await readInput(p),src=await loadPdf(file,buffer),total=src.getPageCount();
  let groups=[];if(String(ranges||'').trim()){for(const part of String(ranges).replace(/[，、;；\s]+/g,',').split(',').filter(Boolean))groups.push(parsePages(part,total));}
  else{const n=Math.floor(Number(every));if(!(n>=1&&n<=1000))throw Error('“每几页一个文件”请填 1 以上的数字。');for(let i=1;i<=total;i+=n)groups.push(Array.from({length:Math.min(n,total-i+1)},(_,k)=>i+k));}
  if(groups.length<2&&groups[0]?.length===total)throw Error('按这个设置只会得到一个和原文件一样的 PDF，请改小页数或分段。');if(groups.length>500)throw Error('最多拆成 500 个文件。');
  await fs.mkdir(outputDir,{recursive:true});const base=safe(path.basename(file,path.extname(file))),folder=await freeName(outputDir,base+'-拆分','');await fs.mkdir(folder);
  const {PDFDocument}=pdfLib(),width=String(groups.length).length,files=[];
  for(let g=0;g<groups.length;g++){const out=await PDFDocument.create();(await out.copyPages(src,groups[g].map(n=>n-1))).forEach(pg=>out.addPage(pg));
    const name=String(g+1).padStart(Math.max(2,width),'0')+'-第'+rangeLabel(groups[g])+'页.pdf';await fs.writeFile(path.join(folder,name),await savePdf(out));files.push(name);}
  return {folder,count:files.length,files,total};}

// 图片转 PDF：每张图一页。size：A4（图片缩放放进 A4，横图自动横放）| fit（页面和图片一样大）。
export async function imagesToPdf({paths,size='A4',outputDir}){const list=[].concat(paths||[]).map(p=>String(p).trim()).filter(Boolean);if(!list.length)throw Error('请选择要放进 PDF 的图片（JPG 或 PNG）。');if(list.length>LIMITS.images)throw Error('一次最多 '+LIMITS.images+' 张图片。');
  const {PDFDocument,PageSizes}=pdfLib(),out=await PDFDocument.create(),MARGIN=28;let bytes=0;
  for(const p of list){const {file,buffer}=await readInput(p,{exts:['.jpg','.jpeg','.png'],limitMB:LIMITS.imageMB});bytes+=buffer.length;if(bytes>LIMITS.totalMB*1024*1024)throw Error('图片加起来超过 '+LIMITS.totalMB+' MB，请分几次。');
    let img;const png=buffer[0]===0x89&&buffer[1]===0x50,jpg=buffer[0]===0xff&&buffer[1]===0xd8;try{img=png?await out.embedPng(buffer):jpg?await out.embedJpg(buffer):null;}catch{img=null;}
    if(!img)throw Error('“'+path.basename(file)+'”读不出来。只支持 JPG、PNG 图片；其他格式请先用画图等软件另存为 JPG。');
    if(size==='fit'){const pg=out.addPage([img.width*0.75,img.height*0.75]);pg.drawImage(img,{x:0,y:0,width:pg.getWidth(),height:pg.getHeight()});continue;}
    const [w0,h0]=PageSizes.A4,land=img.width>img.height,[W,H]=land?[h0,w0]:[w0,h0],pg=out.addPage([W,H]);const k=Math.min((W-2*MARGIN)/img.width,(H-2*MARGIN)/img.height,1.5);
    const w=img.width*k,h=img.height*k;pg.drawImage(img,{x:(W-w)/2,y:(H-h)/2,width:w,height:h});}
  await fs.mkdir(outputDir,{recursive:true});const first=safe(path.basename(list[0],path.extname(list[0])));const target=await freeName(outputDir,list.length>1?first+'-等 '+list.length+' 张图片':first,'.pdf');
  await fs.writeFile(target,await savePdf(out));return {path:target,pages:list.length};}

// PDF 转图片：每页一张 PNG，放进新文件夹。dpi：150（一般）或 200（清晰）。
export async function pdfToImages({path:p,pages='',dpi=150,outputDir,parser}){const {file,buffer}=await readInput(p),src=await loadPdf(file,buffer),total=src.getPageCount();
  const chosen=String(pages||'').trim()?parsePages(pages,total):Array.from({length:total},(_,i)=>i+1);if(chosen.length>LIMITS.renderPages)throw Error('一次最多转 '+LIMITS.renderPages+' 页，请填写页码分几次转。');
  const d=[150,200].includes(Number(dpi))?Number(dpi):150;await ensurePdfEnv();const P=parser||(await import('pdf-parse')).PDFParse,pdf=new P({data:new Uint8Array(buffer)});let shots;
  try{shots=await pdf.getScreenshot({scale:d/72,imageBuffer:true,partial:chosen});}catch(e){throw Error('这台电脑上没法把 PDF 画成图片（'+String(e&&e.message||e).slice(0,80)+'）。');}finally{await pdf.destroy?.();}
  const list=(shots&&shots.pages)||[];if(!list.length)throw Error('没有可转换的页面。');
  await fs.mkdir(outputDir,{recursive:true});const base=safe(path.basename(file,path.extname(file))),folder=await freeName(outputDir,base+'-图片','');await fs.mkdir(folder);
  const width=Math.max(2,String(total).length),files=[];for(let i=0;i<list.length;i++){const n=list[i].pageNumber||chosen[i]||i+1,name=base+'-'+String(n).padStart(width,'0')+'.png';await fs.writeFile(path.join(folder,name),Buffer.from(list[i].data));files.push(name);}
  return {folder,count:files.length,files,dpi:d};}

// 页数（面板里选好文件后显示）。
export async function pdfInfo({path:p}){const {file,buffer}=await readInput(p),src=await loadPdf(file,buffer);return {name:path.basename(file),pages:src.getPageCount(),size:buffer.length};}
