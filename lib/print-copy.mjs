// All transformations return a new buffer. The source file is never written.
import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import PDFDocument from 'pdfkit';
import {ensurePdfEnv} from './pdf-env.mjs';
async function canvasApi(){try{return await import('@napi-rs/canvas');}catch{throw Error('本机缺少图片处理组件，此功能暂不可用。');}}
export function monochromePixels(data,{threshold=180,binary=false}={}){for(let i=0;i<data.length;i+=4){const a=data[i+3]/255,l=(.2126*data[i]+.7152*data[i+1]+.0722*data[i+2])*a+255*(1-a),v=binary?(l<threshold?0:255):Math.round(l);data[i]=data[i+1]=data[i+2]=v;data[i+3]=255;}return data;}
async function imageCopy(buffer,binary){const {createCanvas,loadImage}=await canvasApi(),image=await loadImage(buffer);if(image.width*image.height>24000000)throw Error('图片尺寸过大，请先压缩。');const c=createCanvas(image.width,image.height),ctx=c.getContext('2d');ctx.drawImage(image,0,0);const pixels=ctx.getImageData(0,0,c.width,c.height);monochromePixels(pixels.data,{binary});ctx.putImageData(pixels,0,0);return c.toBuffer('image/png');}
export function optimizeXml(xml,ext,fax){
 if(ext==='.docx'){
  xml=xml.replace(/<w:color\b[^>]*\/>/g,'<w:color w:val="000000"/>').replace(/<w:shd\b[^>]*\/>/g,'<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>').replace(/<w:highlight\b[^>]*\/>/g,'');
  if(fax)xml=xml.replace(/<w:sz(Cs)?\b[^>]*w:val="(\d+)"[^>]*\/>/g,(m,cs,n)=>Number(n)<21?'<w:sz'+(cs||'')+' w:val="21"/>':m).replace(/(<w:(?:top|left|bottom|right|insideH|insideV)\b[^>]*w:sz=")(\d+)(")/g,(m,a,n,b)=>a+Math.max(8,Number(n))+b);
 }else if(ext==='.xlsx'){
  xml=xml.replace(/<color\b[^>]*\/>/g,'<color rgb="FF000000"/>').replace(/<fgColor\b[^>]*\/>/g,'<fgColor rgb="FFFFFFFF"/>').replace(/<bgColor\b[^>]*\/>/g,'<bgColor rgb="FFFFFFFF"/>');
  if(fax)xml=xml.replace(/<sz\b[^>]*val="([\d.]+)"[^>]*\/>/g,(m,n)=>Number(n)<10.5?'<sz val="10.5"/>':m).replace(/style="(?:hair|thin|dotted)"/g,'style="medium"');
 }else{
  // Backgrounds and shape fills become white; text runs become black.
  xml=xml.replace(/<a:(?:solidFill|gradFill)\b[^>]*>[\s\S]*?<\/a:(?:solidFill|gradFill)>/g,'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>');
  // Runs without direct formatting can inherit a white theme font. Give every
  // text run its own black fill while retaining inherited size and typeface.
  xml=xml.replace(/<a:(rPr|defRPr|endParaRPr)\b([^>]*?)\/>/g,'<a:$1$2><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:$1>');
  xml=xml.replace(/(<a:(?:r|fld)\b[^>]*>)(?!\s*<a:rPr\b)/g,'$1<a:rPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr>');
  xml=xml.replace(/(<a:(?:rPr|defRPr|endParaRPr)\b[^>]*>)([\s\S]*?)(<\/a:(?:rPr|defRPr|endParaRPr)>)/g,(_,a,b,c)=>a+b.replace(/<a:solidFill>.*?<\/a:solidFill>/g,'')+'<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'+c);
  xml=xml.replace(/(<a:ln\b[^>]*>)([\s\S]*?)(<\/a:ln>)/g,(_,a,b,c)=>a+b.replace(/<a:solidFill>.*?<\/a:solidFill>/g,'<a:solidFill><a:srgbClr val="000000"/></a:solidFill>')+c);
  if(fax)xml=xml.replace(/(\bsz=")(\d+)(")/g,(m,a,n,b)=>a+Math.max(1050,Number(n))+b).replace(/(<a:ln\b[^>]*\bw=")(\d+)(")/g,(m,a,n,b)=>a+Math.max(12700,Number(n))+b);
 }
 return xml;
}
async function readInput(file,extensions){if(!path.isAbsolute(file)||!extensions.includes(path.extname(file).toLowerCase()))throw Error('请选择支持的本机文件。');const st=await fs.stat(file);if(st.size>80*1024*1024)throw Error('文件超过 80 MB，请先拆分。');return fs.readFile(file);}
export async function printCopy(input){const file=String(input.path||''),ext=path.extname(file).toLowerCase(),fax=input.mode==='fax';const data=await readInput(file,['.docx','.xlsx','.pptx']),zip=await JSZip.loadAsync(data);let expanded=0;for(const [name,entry] of Object.entries(zip.files)){if(entry.dir)continue;const declared=entry._data?.uncompressedSize||0;expanded+=declared;if(expanded>200*1024*1024)throw Error('文件展开后过大，请先拆分。');if(/^(word|xl|ppt)\/.*\.xml$/.test(name)){zip.file(name,optimizeXml(await entry.async('string'),ext,fax));}else if(/^(word|xl|ppt)\/media\/.*\.(png|jpe?g)$/i.test(name)){
  // Preserve media format and relationship names.
  const {createCanvas,loadImage}=await canvasApi(),im=await loadImage(await imageCopy(await entry.async('nodebuffer'),fax)),c=createCanvas(im.width,im.height);c.getContext('2d').drawImage(im,0,0);zip.file(name,c.toBuffer(/\.png$/i.test(name)?'image/png':'image/jpeg'));
 }}return {name:path.basename(file,ext)+(fax?'-传真优化':'-黑白副本')+ext,buffer:await zip.generateAsync({type:'nodebuffer'})};}
export async function pdfPrintCopy(input){const file=String(input.path||''),data=await readInput(file,['.pdf']);await ensurePdfEnv();const {PDFParse}=await import('pdf-parse'),parser=new PDFParse({data:new Uint8Array(data)});let pages;try{const info=await parser.getInfo();if(info.total>60)throw Error('一次最多处理 60 页。');pages=(await parser.getScreenshot({scale:2,imageBuffer:true,last:60})).pages;}finally{await parser.destroy();}
 const seam=input.mode==='seam';let stamp;if(seam){const stampPath=String(input.stampPath||'');stamp=await readInput(stampPath,['.png']);}
 const doc=new PDFDocument({autoFirstPage:false}),chunks=[],done=new Promise((resolve,reject)=>{doc.on('data',c=>chunks.push(c));doc.on('end',()=>resolve(Buffer.concat(chunks)));doc.on('error',reject);});
 for(let i=0;i<pages.length;i++){const p=pages[i],buffer=Buffer.from(p.data),{loadImage,createCanvas}=await canvasApi(),im=await loadImage(buffer),w=im.width/2,h=im.height/2;doc.addPage({size:[w,h],margin:0});doc.image(seam?buffer:await imageCopy(buffer,true),0,0,{width:w,height:h});if(seam){const st=await loadImage(stamp);if(st.width*st.height>4000000)throw Error('印章图片过大。');const sw=st.width/pages.length,c=createCanvas(Math.max(1,Math.ceil(sw)),st.height);c.getContext('2d').drawImage(st,i*sw,0,sw,st.height,0,0,c.width,c.height);const stampHeight=100,stripWidth=stampHeight*sw/st.height;doc.image(c.toBuffer('image/png'),w-stripWidth,h/2-stampHeight/2,{width:stripWidth,height:stampHeight});}}
 doc.end();return {name:path.basename(file,'.pdf')+(seam?'-骑缝章':'-传真黑白')+'.pdf',buffer:await done};
}
