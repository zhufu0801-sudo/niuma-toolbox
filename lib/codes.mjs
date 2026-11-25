// 办公工具箱 · 二维码和条形码（2026-11-24）。不依赖外部库：二维码按 ISO/IEC 18004（字节模式、纠错等级 M、版本 1～15），
// 条形码用 Code 128（B 字符集，数字多时自动用 C 字符集压缩），输出黑白 PNG。
import zlib from 'node:zlib';
import fs from 'node:fs/promises';
import path from 'node:path';

// ---- 二维码 ----
// 纠错等级 M：每个版本的 [第一组块数, 每块总码字, 每块数据码字, 第二组块数, 每块总码字, 每块数据码字]
const RS_M=[null,[1,26,16],[1,44,28],[1,70,44],[2,50,32],[2,67,43],[4,43,27],[4,49,31],[2,60,38,2,61,39],[3,58,36,2,59,37],[4,69,43,1,70,44],[1,80,50,4,81,51],[6,58,36,2,59,37],[8,59,37,1,60,38],[4,64,40,5,65,41],[5,65,41,5,66,42]];
const ALIGN=[null,[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70]];
const REMAINDER=[0,0,7,7,7,7,7,0,0,0,0,0,0,0,3,3];
const EXP=new Array(512),LOG=new Array(256);
(()=>{let x=1;for(let i=0;i<255;i++){EXP[i]=x;LOG[x]=i;x<<=1;if(x&256)x^=0x11d;}for(let i=255;i<512;i++)EXP[i]=EXP[i-255];})();
function gmul(a,b){return a&&b?EXP[LOG[a]+LOG[b]]:0;}
function rsGen(n){let g=[1];for(let i=0;i<n;i++){const next=new Array(g.length+1).fill(0);for(let j=0;j<g.length;j++){next[j]^=g[j];next[j+1]^=gmul(g[j],EXP[i]);}g=next;}return g;}
function rsEncode(data,n){const g=rsGen(n),res=data.concat(new Array(n).fill(0));for(let i=0;i<data.length;i++){const c=res[i];if(c)for(let j=0;j<g.length;j++)res[i+j]^=gmul(g[j],c);}return res.slice(data.length);}
function capacity(v){const t=RS_M[v];return t[0]*t[2]+(t[3]||0)*(t[5]||0);}
function pickVersion(len){for(let v=1;v<RS_M.length;v++){const bits=4+(v<10?8:16)+8*len;if(bits<=capacity(v)*8)return v;}return 0;}
export const QR_MAX_BYTES=capacity(15)-3;
function dataCodewords(bytes,v){const bits=[],put=(val,n)=>{for(let i=n-1;i>=0;i--)bits.push((val>>>i)&1);};
  put(4,4);put(bytes.length,v<10?8:16);for(const b of bytes)put(b,8);const total=capacity(v)*8;for(let i=0;i<4&&bits.length<total;i++)bits.push(0);while(bits.length%8)bits.push(0);
  const out=[];for(let i=0;i<bits.length;i+=8){let x=0;for(let j=0;j<8;j++)x=(x<<1)|bits[i+j];out.push(x);}for(let pad=0;out.length<capacity(v);pad++)out.push(pad%2?0x11:0xec);return out;}
function interleave(data,v){const t=RS_M[v],blocks=[];let k=0;const groups=[[t[0],t[1],t[2]]];if(t[3])groups.push([t[3],t[4],t[5]]);
  for(const [count,total,dc] of groups)for(let i=0;i<count;i++){const d=data.slice(k,k+dc);k+=dc;blocks.push({d,e:rsEncode(d,total-dc)});}
  const out=[],maxD=Math.max(...blocks.map(b=>b.d.length)),maxE=Math.max(...blocks.map(b=>b.e.length));
  for(let i=0;i<maxD;i++)for(const b of blocks)if(i<b.d.length)out.push(b.d[i]);for(let i=0;i<maxE;i++)for(const b of blocks)if(i<b.e.length)out.push(b.e[i]);return out;}
function bch(value,poly,bits){let v=value<<bits;const top=Math.floor(Math.log2(poly));for(let i=Math.floor(Math.log2(v||1));i>=top;i--)if(v&(1<<i))v^=poly<<(i-top);return (value<<bits)|v;}
const MASKS=[(r,c)=>(r+c)%2===0,(r)=>r%2===0,(r,c)=>c%3===0,(r,c)=>(r+c)%3===0,(r,c)=>(Math.floor(r/2)+Math.floor(c/3))%2===0,(r,c)=>(r*c)%2+(r*c)%3===0,(r,c)=>((r*c)%2+(r*c)%3)%2===0,(r,c)=>((r+c)%2+(r*c)%3)%2===0];
function baseMatrix(v){const n=17+4*v,m=Array.from({length:n},()=>new Array(n).fill(null));
  const finder=(r0,c0)=>{for(let r=-1;r<=7;r++)for(let c=-1;c<=7;c++){const rr=r0+r,cc=c0+c;if(rr<0||cc<0||rr>=n||cc>=n)continue;const on=r>=0&&r<=6&&c>=0&&c<=6&&(r===0||r===6||c===0||c===6||(r>=2&&r<=4&&c>=2&&c<=4));m[rr][cc]=on;}};
  finder(0,0);finder(0,n-7);finder(n-7,0);
  for(let i=8;i<n-8;i++){m[6][i]=i%2===0;m[i][6]=i%2===0;}
  const a=ALIGN[v],last=a[a.length-1];for(const r of a)for(const c of a){if((r===6&&c===6)||(r===6&&c===last)||(r===last&&c===6))continue;/* 和定位图形重叠的不放 */for(let dr=-2;dr<=2;dr++)for(let dc=-2;dc<=2;dc++)m[r+dr][c+dc]=Math.max(Math.abs(dr),Math.abs(dc))!==1;}
  m[n-8][8]=true;// 暗模块
  for(let i=0;i<9;i++){if(m[8][i]===null)m[8][i]=false;if(m[i][8]===null)m[i][8]=false;}for(let i=0;i<8;i++){if(m[8][n-1-i]===null)m[8][n-1-i]=false;if(m[n-1-i][8]===null)m[n-1-i][8]=false;}
  if(v>=7){const info=bch(v,0x1f25,12);for(let i=0;i<18;i++){const bit=((info>>i)&1)===1,r=Math.floor(i/3),c=n-11+i%3;m[r][c]=bit;m[c][r]=bit;}}
  return m;}
function placeData(m,codewords,v){const n=m.length,fixed=m.map(row=>row.map(x=>x!==null)),bits=[];for(const cw of codewords)for(let i=7;i>=0;i--)bits.push((cw>>i)&1);for(let i=0;i<REMAINDER[v];i++)bits.push(0);
  let k=0,up=true;for(let col=n-1;col>0;col-=2){if(col===6)col--;for(let i=0;i<n;i++){const r=up?n-1-i:i;for(let j=0;j<2;j++){const c=col-j;if(fixed[r][c])continue;m[r][c]=k<bits.length?bits[k]===1:false;k++;}}up=!up;}return fixed;}
function formatBits(mask){return bch((0<<3)|mask,0x537,10)^0x5412;}// 纠错等级 M = 00
function applyFormat(m,mask){const n=m.length,f=formatBits(mask);for(let i=0;i<15;i++){const bit=((f>>i)&1)===1;
    if(i<6)m[i][8]=bit;else if(i<8)m[i+1][8]=bit;else m[n-15+i][8]=bit;
    if(i<8)m[8][n-1-i]=bit;else if(i<9)m[8][15-i-1+1]=bit;else m[8][15-i-1]=bit;}m[n-8][8]=true;}
function penalty(m){const n=m.length;let p=0;
  for(let r=0;r<n;r++){let run=1;for(let c=1;c<n;c++){if(m[r][c]===m[r][c-1]){run++;if(run===5)p+=3;else if(run>5)p++;}else run=1;}}
  for(let c=0;c<n;c++){let run=1;for(let r=1;r<n;r++){if(m[r][c]===m[r-1][c]){run++;if(run===5)p+=3;else if(run>5)p++;}else run=1;}}
  for(let r=0;r<n-1;r++)for(let c=0;c<n-1;c++){const x=m[r][c];if(x===m[r][c+1]&&x===m[r+1][c]&&x===m[r+1][c+1])p+=3;}
  const pat=[true,false,true,true,true,false,true],w=[false,false,false,false];
  const line=(get)=>{for(let i=0;i+7<=n;i++){let hit=true;for(let j=0;j<7;j++)if(get(i+j)!==pat[j]){hit=false;break;}if(!hit)continue;let before=i>=4,after=i+11<=n;for(let j=1;j<=4&&before;j++)if(get(i-j))before=false;for(let j=7;j<11&&after;j++)if(get(i+j))after=false;if(before||after)p+=40;}};
  for(let r=0;r<n;r++)line(c=>m[r][c]);for(let c=0;c<n;c++)line(r=>m[r][c]);
  let dark=0;for(const row of m)for(const x of row)if(x)dark++;p+=Math.floor(Math.abs(dark*20-n*n*10)/(n*n))*10;void w;return p;}
export function qrMatrix(text){const bytes=[...Buffer.from(String(text),'utf8')];if(!bytes.length)throw Error('内容是空的。');const v=pickVersion(bytes.length);if(!v)throw Error('内容太长，二维码最多约 '+QR_MAX_BYTES+' 个英文字母或数字（中文约 '+Math.floor(QR_MAX_BYTES/3)+' 个字）。');
  const code=interleave(dataCodewords(bytes,v),v);let best=null,bestScore=Infinity;
  for(let mask=0;mask<8;mask++){const m=baseMatrix(v),fixed=placeData(m,code,v);for(let r=0;r<m.length;r++)for(let c=0;c<m.length;c++)if(!fixed[r][c]&&MASKS[mask](r,c))m[r][c]=!m[r][c];applyFormat(m,mask);const s=penalty(m);if(s<bestScore){bestScore=s;best=m;}}
  return best;}

// ---- Code 128 ----
const C128=['212222','222122','222221','121223','121322','131222','122213','122312','132212','221213','221312','231212','112232','122132','122231','113222','123122','123221','223211','221132','221231','213212','223112','312131','311222','321122','321221','312212','322112','322211','212123','212321','232121','111323','131123','131321','112313','132113','132311','211313','231113','231311','112133','112331','132131','113123','113321','133121','313121','211331','231131','213113','213311','213131','311123','311321','331121','312113','312311','332111','314111','221411','431111','111224','111422','121124','121421','141122','141221','112214','112412','122114','122411','142112','142211','241211','221114','413111','241112','134111','111242','121142','121241','114212','124112','124211','411212','421112','421211','212141','214121','412121','111143','111341','131141','114113','114311','411113','411311','113141','114131','311141','411131','211412','211214','211232','2331112'];
export function code128Values(text){const s=String(text);if(!s)throw Error('内容是空的。');if(s.length>80)throw Error('条形码内容最多 80 个字符。');if(/[^\x20-\x7e]/.test(s))throw Error('条形码只能用英文字母、数字和常用符号；中文请用二维码。');
  const vals=[];let set='';const digits=(i)=>{let k=i;while(k<s.length&&/\d/.test(s[k]))k++;return k-i;};
  let i=0;while(i<s.length){const d=digits(i);if(d>=4&&(set==='C'||d%2===0||d>=6)){const run=d-d%2;if(set!=='C'){vals.push(set?99:105);set='C';}for(let j=0;j<run;j+=2)vals.push(+s.substr(i+j,2));i+=run;continue;}
    if(set!=='B'){vals.push(set?100:104);set='B';}vals.push(s.charCodeAt(i)-32);i++;}
  let sum=vals[0];for(let k=1;k<vals.length;k++)sum+=vals[k]*k;vals.push(sum%103);vals.push(106);return vals;}
export function code128Modules(text){const out=[];for(const v of code128Values(text)){const p=C128[v];let bar=true;for(const ch of p){for(let k=0;k<+ch;k++)out.push(bar);bar=!bar;}}return out;}

// ---- PNG（灰度 8 位）----
const CRC=new Int32Array(256).map((_,n)=>{let c=n;for(let k=0;k<8;k++)c=c&1?0xedb88320^(c>>>1):c>>>1;return c;});
function crc32(buf){let c=-1;for(const b of buf)c=CRC[(c^b)&255]^(c>>>8);return (c^-1)>>>0;}
function chunk(type,data){const len=Buffer.alloc(4);len.writeUInt32BE(data.length);const td=Buffer.concat([Buffer.from(type,'latin1'),data]),crc=Buffer.alloc(4);crc.writeUInt32BE(crc32(td));return Buffer.concat([len,td,crc]);}
export function png(width,height,pixel){const raw=Buffer.alloc((width+1)*height);for(let y=0;y<height;y++){raw[y*(width+1)]=0;for(let x=0;x<width;x++)raw[y*(width+1)+1+x]=pixel(x,y)?0:255;}
  const ihdr=Buffer.alloc(13);ihdr.writeUInt32BE(width,0);ihdr.writeUInt32BE(height,4);ihdr[8]=8;ihdr[9]=0;ihdr[10]=0;ihdr[11]=0;ihdr[12]=0;
  return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]),chunk('IHDR',ihdr),chunk('IDAT',zlib.deflateSync(raw)),chunk('IEND',Buffer.alloc(0))]);}
export function qrPng(text,scale=8){const m=qrMatrix(text),q=4,n=m.length,size=(n+2*q)*scale;return png(size,size,(x,y)=>{const r=Math.floor(y/scale)-q,c=Math.floor(x/scale)-q;return r>=0&&c>=0&&r<n&&c<n&&m[r][c];});}
export function barcodePng(text,scale=2,height=90){const mods=code128Modules(text),q=10,w=(mods.length+2*q)*scale;return png(w,height,(x)=>{const k=Math.floor(x/scale)-q;return k>=0&&k<mods.length&&mods[k];});}

// ---- 批量生成到文件夹 ----
function safeName(s){return String(s).replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').replace(/\s+/g,' ').trim().slice(0,40)||'code';}
export async function makeCodes({items,kind,outputDir}){const list=[].concat(items||[]).map(x=>String(x==null?'':x));if(!list.length)throw Error('没有要生成的内容。');if(list.length>500)throw Error('一次最多生成 500 个。');
  const qr=kind!=='barcode',stamp=new Date(),pad=n=>String(n).padStart(2,'0'),dir=path.join(outputDir,(qr?'二维码':'条形码')+'-'+stamp.getFullYear()+pad(stamp.getMonth()+1)+pad(stamp.getDate())+'-'+pad(stamp.getHours())+pad(stamp.getMinutes())+pad(stamp.getSeconds()));
  await fs.mkdir(dir,{recursive:true});const files=[],errors=[];
  for(let i=0;i<list.length;i++){const text=list[i].trim();if(!text){files.push('');continue;}try{const data=qr?qrPng(text):barcodePng(text),file=path.join(dir,String(i+1).padStart(3,'0')+'-'+safeName(text)+'.png');await fs.writeFile(file,data);files.push(file);}catch(e){files.push('');errors.push({index:i,text:text.slice(0,30),error:e.message});}}
  return {folder:dir,files,errors};}
