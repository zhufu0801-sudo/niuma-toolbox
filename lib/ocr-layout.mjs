// 扫描件版面还原：用文字识别给出的每个词的位置，按“行”重新排顺序、找出表格、标出居中标题和靠右的落款。
// 2026-11-20 体验测试：Windows 文字识别按区域输出，表格被按列竖着排、同一行的“申领人：部门：”被拆开。
import {tidyOcr} from './doc-tools.mjs';

const median=a=>{const s=a.slice().sort((x,y)=>x-y);return s.length?s[Math.floor(s.length/2)]:0;};
const CJK=/[぀-ヿ㐀-鿿豈-﫿＀-￯]/;

// lines：[{t, words:[{t,x,y,w,h}]}]，坐标是图片像素；width：图片宽度。
export function layoutOcr(lines,width){
  const words=[];for(const l of lines||[])for(const w of l.words||[])if(w&&String(w.t||'').trim()&&w.w>0&&w.h>0)words.push({t:String(w.t),x:+w.x,y:+w.y,w:+w.w,h:+w.h,line:l});
  if(words.length<3)return null;
  const H=median(words.map(w=>w.h))||10,W=width||Math.max(...words.map(w=>w.x+w.w));
  // 1) 同一识别行里，词与词之间空得很开（> 1.5 个字高）就拆成不同的“块”（多半是表格的不同格子）。
  const chunks=[];
  for(const l of lines||[]){const ws=words.filter(w=>w.line===l).sort((a,b)=>a.x-b.x);let cur=null;
    for(const w of ws){if(cur&&w.x-(cur.x1)<=1.5*H){cur.parts.push(w);cur.x1=Math.max(cur.x1,w.x+w.w);cur.y0=Math.min(cur.y0,w.y);cur.y1=Math.max(cur.y1,w.y+w.h);}
      else{cur={parts:[w],x0:w.x,x1:w.x+w.w,y0:w.y,y1:w.y+w.h};chunks.push(cur);}}}
  // 词之间按实际距离决定要不要空格：挨着的（中文、数字逐字识别时）直接连上，有明显空隙的英文单词之间留一个空格。
  for(const c of chunks){c.t=tidyOcr(c.parts.map((p,i)=>{if(!i)return p.t;const prev=c.parts[i-1],gap=p.x-(prev.x+prev.w);return (gap>0.3*H&&!(CJK.test(prev.t.slice(-1))&&CJK.test(p.t[0]))?' ':'')+p.t;}).join(''));c.cy=(c.y0+c.y1)/2;}
  // 2) 按垂直位置分行：中心线相差不到半个字高的算同一行。
  chunks.sort((a,b)=>a.cy-b.cy);const rows=[];
  for(const c of chunks){const r=rows[rows.length-1];if(r&&Math.abs(c.cy-r.cy)<0.5*H){r.items.push(c);r.cy=(r.cy*(r.items.length-1)+c.cy)/r.items.length;r.y1=Math.max(r.y1,c.y1);}else rows.push({items:[c],cy:c.cy,y0:c.y0,y1:c.y1});}
  for(const r of rows){r.items.sort((a,b)=>a.x0-b.x0);r.x0=r.items[0].x0;r.x1=Math.max(...r.items.map(c=>c.x1));}
  // 3) 连续 2 行以上、每行至少 3 块、行距不大的，当作表格；按块在横向上的覆盖范围合并出列。
  const isTableRow=r=>r.items.length>=3;const segs=[];let i=0;
  const flushText=buf=>{if(buf.length)segs.push({text:buf.join('\n')});buf.length=0;};const buf=[];
  while(i<rows.length){
    if(isTableRow(rows[i])){let j=i;while(j+1<rows.length&&isTableRow(rows[j+1])&&rows[j+1].y0-rows[j].y1<2.5*H)j++;
      if(j>i){flushText(buf);segs.push({table:toTable(rows.slice(i,j+1),H)});i=j+1;continue;}}
    const r=rows[i],para=special(r,W,H,i);
    if(para){flushText(buf);segs.push({para});}else buf.push(joinRow(r,H));
    i++;}
  flushText(buf);
  return segs;
}
function joinRow(r,H){let out='';r.items.forEach((c,k)=>{const gap=k?c.x0-r.items[k-1].x1:0;out+=(k?(gap>3*H?'　　':' '):'')+c.t;});return out;}
// 居中的短行（标题）、靠右的行（落款、日期）、带“：”的填写栏单独成段，不和上下文合并。
function special(r,W,H,index){
  const text=formRow(r,H),width=r.x1-r.x0,center=(r.x0+r.x1)/2;
  if(index<3&&width<W*0.7&&Math.abs(center-W/2)<W*0.08&&r.items.length===1&&text.length<=40)return {text,level:1};
  if(r.x0>W*0.55&&text.length<=30)return {text,align:'right'};
  if(text!==joinRow(r,H))return {text};
  return null;
}
// 填写栏：“申领人：”后面原本是下划线（识别不到），补上“＿＿＿＿＿＿”留空。
// 只有一行里有两个以上“xx：”栏目时才补（单独一行的“各部门：”是称呼，不补）。
function formRow(r,H){const fields=r.items.filter(c=>/[:：]$/.test(c.t)).length;if(fields<2&&!(fields===1&&r.items.length>1&&/[:：]$/.test(r.items[0].t)&&r.items[1].x0-r.items[0].x1>2*H))return joinRow(r,H);let out='';r.items.forEach((c,k)=>{const gap=k?c.x0-r.items[k-1].x1:0;out+=(k?(gap>3*H?'　　':' '):'')+c.t;const next=r.items[k+1];if(/[:：]$/.test(c.t)&&(!next||next.x0-c.x1>2*H))out+='＿＿＿＿＿＿';});return out;}
function toTable(rows,H){
  const spans=[];for(const r of rows)for(const c of r.items)spans.push([c.x0,c.x1]);spans.sort((a,b)=>a[0]-b[0]);
  const cols=[];for(const s of spans){const last=cols[cols.length-1];if(last&&s[0]<=last[1]+0.3*H)last[1]=Math.max(last[1],s[1]);else cols.push([s[0],s[1]]);}
  // 居中的表头和靠左的数据可能各自成一列：列数多于“一行里最多的格子数”时，把挨得最近的两列合并。
  const want=Math.max(...rows.map(r=>r.items.length));while(cols.length>want){let k=0,g=Infinity;for(let i=0;i+1<cols.length;i++){const d=cols[i+1][0]-cols[i][1];if(d<g){g=d;k=i;}}cols.splice(k,2,[cols[k][0],Math.max(cols[k][1],cols[k+1][1])]);}
  return rows.map(r=>{const cells=cols.map(()=>[]);for(const c of r.items){let best=0,score=-1;cols.forEach((col,k)=>{const o=Math.min(col[1],c.x1)-Math.max(col[0],c.x0);if(o>score){score=o;best=k;}});cells[best].push(c.t);}return cells.map(x=>x.join(' '));});
}
