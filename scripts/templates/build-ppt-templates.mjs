import PptxGenJS from 'pptxgenjs';
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {deflateSync} from 'node:zlib';
import {addMotion} from '../../lib/pptx-motion.mjs';
import {pptThemes as themes} from '../../lib/ppt-themes.mjs';
import {magazineSlides} from './ppt-magazine.mjs';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'../..');
const out=path.join(root,'public/templates/ppt');await fs.mkdir(out,{recursive:true});
// All twelve scenes share editable layout primitives.
function pngGradient(a,b,w=960,h=540){const rgb=x=>x.match(/../g).map(x=>parseInt(x,16)),c=rgb(a),d=rgb(b),raw=Buffer.alloc((w*3+1)*h);for(let y=0;y<h;y++)for(let x=0;x<w;x++){let t=Math.min(1,Math.max(0,Math.hypot((x-w*.9)/w,(y-h*.3)/h)));let p=y*(w*3+1)+1+x*3;for(let k=0;k<3;k++)raw[p+k]=Math.round(c[k]*t+d[k]*(1-t));}function chunk(type,data){let s=Buffer.concat([Buffer.from(type),data]),crc=0xffffffff;for(const byte of s){crc^=byte;for(let j=0;j<8;j++)crc=(crc>>>1)^((crc&1)?0xedb88320:0);}let len=Buffer.alloc(4),end=Buffer.alloc(4);len.writeUInt32BE(data.length);end.writeUInt32BE((crc^0xffffffff)>>>0);return Buffer.concat([len,s,end]);}let head=Buffer.alloc(13);head.writeUInt32BE(w);head.writeUInt32BE(h,4);head[8]=8;head[9]=2;return Buffer.concat([Buffer.from('89504e470d0a1a0a','hex'),chunk('IHDR',head),chunk('IDAT',deflateSync(raw)),chunk('IEND',Buffer.alloc(0))]);}
const layouts=['cover','agenda','chapter','points','metric','timeline','compare','picture','quote','closing'];
const index=[];
for(const t of themes)for(const lang of ['zh','ja']){
 const ja=lang==='ja',font=ja?'Yu Gothic':'Microsoft YaHei',pick=(zh,jp)=>ja?jp:zh,id=t.id+'-'+lang;
 const ppt=new PptxGenJS();ppt.layout='LAYOUT_WIDE';ppt.author='Office Toolbox';ppt.subject='Editable presentation template';ppt.title=t[lang];ppt.lang=ja?'ja-JP':'zh-CN';ppt.theme={headFontFace:font,bodyFontFace:font,lang:ppt.lang};
 const gradient=t.gradient?pngGradient(t.bg,t.gradientEnd):null;
 const motions=magazineSlides(ppt,t,lang,gradient);
 let buffer=await ppt.write({outputType:'nodebuffer'});buffer=await addMotion(buffer,motions);await fs.writeFile(path.join(out,id+'.pptx'),buffer);
 index.push({id,title:t[lang],lang,file:id+'.pptx',preview:id+'.png',kind:'ppt',theme:t.id,colors:{background:t.bg,accent:t.accent},layouts});
}
await fs.writeFile(path.join(out,'index.json'),JSON.stringify(index,null,2));console.log('PPT 模板：'+index.length+' 份，每份 10 页。');
