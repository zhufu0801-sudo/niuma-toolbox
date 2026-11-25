// 办公工具箱 · PNG 拼长图（2026-11-24）：把 WPS 演示每页导出的 PNG 上下拼成一张长图。只用 Node 自带的 zlib，不需要装图形库。
import zlib from 'node:zlib';
import fs from 'node:fs/promises';
import path from 'node:path';

const SIG=Buffer.from([137,80,78,71,13,10,26,10]);
const CRC=new Int32Array(256).map((_,n)=>{let c=n;for(let k=0;k<8;k++)c=c&1?0xedb88320^(c>>>1):c>>>1;return c;});
function crc32(buf){let c=-1;for(const b of buf)c=CRC[(c^b)&255]^(c>>>8);return (c^-1)>>>0;}
function chunk(type,data){const len=Buffer.alloc(4);len.writeUInt32BE(data.length);const td=Buffer.concat([Buffer.from(type,'latin1'),data]),crc=Buffer.alloc(4);crc.writeUInt32BE(crc32(td));return Buffer.concat([len,td,crc]);}
// 读 8 位、非隔行的灰度 / RGB / RGBA / 调色板 PNG，统一成 RGBA。
export function decodePng(buf){if(!buf.subarray(0,8).equals(SIG))throw Error('不是 PNG 图片。');let pos=8,ihdr=null,plte=null,trns=null;const idat=[];
  while(pos<buf.length){const len=buf.readUInt32BE(pos),type=buf.toString('latin1',pos+4,pos+8),data=buf.subarray(pos+8,pos+8+len);pos+=12+len;
    if(type==='IHDR')ihdr={w:data.readUInt32BE(0),h:data.readUInt32BE(4),depth:data[8],color:data[9],interlace:data[12]};else if(type==='PLTE')plte=data;else if(type==='tRNS')trns=data;else if(type==='IDAT')idat.push(data);else if(type==='IEND')break;}
  if(!ihdr)throw Error('PNG 图片不完整。');if(ihdr.depth!==8||ihdr.interlace)throw Error('暂不支持这种 PNG（'+ihdr.depth+' 位'+(ihdr.interlace?'、隔行':'')+'）。');
  const ch={0:1,2:3,3:1,4:2,6:4}[ihdr.color];if(!ch)throw Error('暂不支持这种 PNG 颜色类型。');const {w,h}=ihdr,stride=w*ch,raw=zlib.inflateSync(Buffer.concat(idat)),px=Buffer.alloc(stride*h);
  for(let y=0;y<h;y++){const f=raw[y*(stride+1)],src=raw.subarray(y*(stride+1)+1,(y+1)*(stride+1)),cur=px.subarray(y*stride,(y+1)*stride),prev=y?px.subarray((y-1)*stride,y*stride):null;
    for(let x=0;x<stride;x++){const a=x>=ch?cur[x-ch]:0,b=prev?prev[x]:0,c=prev&&x>=ch?prev[x-ch]:0;let v=src[x];
      if(f===1)v+=a;else if(f===2)v+=b;else if(f===3)v+=(a+b)>>1;else if(f===4){const p=a+b-c,pa=Math.abs(p-a),pb=Math.abs(p-b),pc=Math.abs(p-c);v+=pa<=pb&&pa<=pc?a:pb<=pc?b:c;}cur[x]=v&255;}}
  const rgba=Buffer.alloc(w*h*4);for(let i=0;i<w*h;i++){let r,g,b,al=255;if(ihdr.color===0){r=g=b=px[i];}else if(ihdr.color===4){r=g=b=px[i*2];al=px[i*2+1];}else if(ihdr.color===2){r=px[i*3];g=px[i*3+1];b=px[i*3+2];}else if(ihdr.color===6){r=px[i*4];g=px[i*4+1];b=px[i*4+2];al=px[i*4+3];}else{const k=px[i];r=plte[k*3];g=plte[k*3+1];b=plte[k*3+2];if(trns&&k<trns.length)al=trns[k];}
    rgba[i*4]=r;rgba[i*4+1]=g;rgba[i*4+2]=b;rgba[i*4+3]=al;}return {width:w,height:h,data:rgba};}
export function encodePng({width,height,data}){const stride=width*4,raw=Buffer.alloc((stride+1)*height);for(let y=0;y<height;y++){raw[y*(stride+1)]=0;data.copy(raw,y*(stride+1)+1,y*stride,(y+1)*stride);}
  const ihdr=Buffer.alloc(13);ihdr.writeUInt32BE(width,0);ihdr.writeUInt32BE(height,4);ihdr[8]=8;ihdr[9]=6;return Buffer.concat([SIG,chunk('IHDR',ihdr),chunk('IDAT',zlib.deflateSync(raw,{level:6})),chunk('IEND',Buffer.alloc(0))]);}
// 等比缩放到指定宽度（最近邻取样，够用来拼预览长图）。
function resize(img,width){if(img.width===width)return img;const k=img.width/width,height=Math.max(1,Math.round(img.height/k)),out=Buffer.alloc(width*height*4);
  for(let y=0;y<height;y++){const sy=Math.min(img.height-1,Math.floor(y*k));for(let x=0;x<width;x++){const sx=Math.min(img.width-1,Math.floor(x*k));img.data.copy(out,(y*width+x)*4,(sy*img.width+sx)*4,(sy*img.width+sx)*4+4);}}return {width,height,data:out};}
export async function stitchVertical({paths,output,gap=16,background=[255,255,255]}){const list=[].concat(paths||[]).map(String);if(list.length<2)throw Error('至少要两张图片才能拼长图。');if(list.length>200)throw Error('一次最多拼 200 张。');
  const imgs=[];for(const p of list){if(!path.isAbsolute(p))throw Error('图片位置无效。');imgs.push(decodePng(await fs.readFile(p)));}
  const width=Math.min(2000,Math.max(...imgs.map(i=>i.width))),scaled=imgs.map(i=>resize(i,width)),height=scaled.reduce((s,i)=>s+i.height,0)+gap*(scaled.length-1);
  if(width*height>120e6)throw Error('长图太大了（超过 1.2 亿像素），请分几次或少选几页。');
  const data=Buffer.alloc(width*height*4);for(let i=0;i<width*height;i++){data[i*4]=background[0];data[i*4+1]=background[1];data[i*4+2]=background[2];data[i*4+3]=255;}
  let top=0;for(const im of scaled){for(let y=0;y<im.height;y++)for(let x=0;x<width;x++){const s=(y*width+x)*4,d=((top+y)*width+x)*4,a=im.data[s+3]/255;data[d]=Math.round(im.data[s]*a+data[d]*(1-a));data[d+1]=Math.round(im.data[s+1]*a+data[d+1]*(1-a));data[d+2]=Math.round(im.data[s+2]*a+data[d+2]*(1-a));}top+=im.height+gap;}
  let target=output,n=1;const ext=path.extname(output),base=output.slice(0,output.length-ext.length);for(;;){try{await fs.writeFile(target,encodePng({width,height,data}),{flag:'wx'});break;}catch(e){if(e.code!=='EEXIST'||n>50)throw e;target=base+'-'+(++n)+ext;}}
  return {path:target,width,height,count:scaled.length};}
