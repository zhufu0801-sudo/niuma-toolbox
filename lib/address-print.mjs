// 宛名印刷：日本はがき / 長形3号（竖排、邮编对准预印格子）、中国 5 号信封（横排、左上邮编格子）、地址标签。
// 只用 docx（已有依赖）。尺寸单位：毫米，写入时换成 twip。
import {Paragraph,TextRun,Table,TableRow,TableCell,WidthType,TextDirection,TableAnchorType,OverlapType,TableLayoutType,HeightRule,BorderStyle,VerticalAlign,AlignmentType,PageOrientation} from 'docx';

const tw=mm=>Math.round(mm*56.6929);
const NONE={style:BorderStyle.NONE,size:0,color:'FFFFFF'};
const NO_BORDERS={top:NONE,bottom:NONE,left:NONE,right:NONE};
const BOX={style:BorderStyle.SINGLE,size:6,color:'C8102E'};
const toHalf=s=>String(s).replace(/[０-９]/g,c=>String.fromCharCode(c.charCodeAt(0)-0xFEE0));

// 纸张（宽 × 高，毫米）与邮编格子位置。日本格子位置按日本郵便的はがき规格；長形3号按右上对齐同样的偏移。
// 实际纸张印刷前建议先用普通纸试印一张对位。
export const PAPERS={
  postcard:{w:100,h:148,style:'ja',zip:{top:12,h:8,w:5.7,x:[44,51,58,65.8,72.6,79.4,86.2]}},
  envelope:{w:120,h:235,style:'ja',zip:{top:12,h:8,w:5.7,x:[64,71,78,85.8,92.6,99.4,106.2]}},
  cn5:{w:220,h:110,style:'zh',zip:{top:10,h:10,w:8,x:[20,30,40,50,60,70]}}
};

// 一行地址 → {zip, address, name}。有 Tab 时按列分（从表格读取的数据），否则按空格分。
// 邮编可在任意位置：日本 〒123-4567 / 1234567，中国 6 位数字。姓名取最后一段。
export function parseAddress(line,style='ja'){
  let raw=toHalf(String(line||'').trim());if(!raw)return null;
  // 从 CSV / 表格逗号复制的行：没有 Tab 时把逗号当作分隔。
  if(!raw.includes('\t')&&/[,，]/.test(raw))raw=raw.split(/\s*[,，]\s*/).join('\t');
  const zipRe=style==='zh'?/(?:邮编|郵編)?[:：]?\s*(?<!\d)(\d{6})(?!\d)/:/〒?\s*(?<!\d)(\d{3})[-－ー‐]?(\d{4})(?!\d)/;
  const tabbed=raw.includes('\t');let fields=(tabbed?raw.split('\t'):raw.split(/[ 　]+/)).map(s=>s.trim()).filter(Boolean);
  let zip='';
  if(tabbed){const i=fields.findIndex(f=>zipRe.test(f)&&f.replace(zipRe,'').trim()==='');if(i>=0){const m=fields[i].match(zipRe);zip=m.slice(1).join('');fields.splice(i,1);}}
  if(!zip){const joined=fields.join(' ');const m=joined.match(zipRe);if(m){zip=m.slice(1).join('');fields=joined.replace(m[0],' ').split(/[ 　]+/).filter(Boolean);}}
  return splitName(fields,style,zip);
}

// 从一行的各段里分出姓名：
// - 单独的敬称（様・御中・殿・收・先生…）并回前一段；
// - 日本人名常写成「鈴木 一郎」：末尾两段都是短的、不含数字和地名字样时，合成一个姓名；
// - 姓名写在最前面（张三 北京市…）时，按“哪一段像地址”来判断。
const HONOR=/^(様|さま|御中|殿|各位|先生|女士|收|启|親展|亲启|親啟)$/;
// 像地址：有数字，或者够长且带行政区划/路名字样（市川、中村、町田这类短姓名不算）。
const addrLike=s=>/[0-9０-９]/.test(s)||(s.length>5&&/[都道府県市区町村郡省路街号楼]|丁目/.test(s));
const shortName=s=>!!s&&s.length<=5&&!addrLike(s)&&!/[都道府県市区郡部課係]$|支店$|営業所$|事業所$/.test(s)&&!/株式会社|有限会社|合同会社|公司|御中|[　 ]/.test(s);
function splitName(fields,style,zip){
  fields=fields.slice();let honor='';
  if(fields.length>1&&HONOR.test(fields[fields.length-1]))honor=fields.pop();
  let name='';const n=fields.length;
  if(n>2&&style!=='zh'&&shortName(fields[n-1])&&shortName(fields[n-2])){const given=fields.pop(),family=fields.pop();name=family+'　'+given;}
  else if(n>1&&addrLike(fields[n-1])&&!addrLike(fields[0]))name=fields.shift();
  else if(n>1)name=fields.pop();
  else if(honor){name=fields.pop()||'';}
  if(honor&&name)name+=(style==='zh'?' ':'　')+honor;else if(honor)fields.push(honor);
  return {zip,address:fields.join(style==='zh'?'':'　'),name};
}

export function withHonorific(name,style){
  name=String(name||'').trim();if(!name)return '';
  if(style==='zh')return /(收|启|啟|鉴|鑒|亲启|親啟)$/.test(name)?name:name+' 收';
  return /(様|さま|御中|殿|各位|先生)$/.test(name)?name:name+'　様';
}

// 竖排时把数字换成汉数字、横线换成长音符、英文字母换成全角，避免横倒。
export function verticalNumbers(s){const K='〇一二三四五六七八九';return toHalf(s).replace(/[0-9]/g,d=>K[d]).replace(/[-－‐−―]/g,'ー').replace(/[A-Za-z]/g,c=>String.fromCharCode(c.charCodeAt(0)+0xFEE0));}

function run(text,size,opt={}){return new TextRun({text,size:size*2,...opt});}
function floatTable({x,y,widths,height,cells}){
  return new Table({layout:TableLayoutType.FIXED,width:{size:tw(widths.reduce((a,b)=>a+b,0)),type:WidthType.DXA},columnWidths:widths.map(tw),borders:NO_BORDERS,
    float:{horizontalAnchor:TableAnchorType.PAGE,verticalAnchor:TableAnchorType.PAGE,absoluteHorizontalPosition:tw(x),absoluteVerticalPosition:tw(y),overlap:OverlapType.OVERLAP},
    rows:[new TableRow({height:{value:tw(height),rule:HeightRule.EXACT},cantSplit:true,children:cells})]});
}
function cell(width,children,opt={}){return new TableCell({width:{size:tw(width),type:WidthType.DXA},margins:{top:0,bottom:0,left:0,right:0},borders:opt.borders||NO_BORDERS,verticalAlign:opt.vAlign||VerticalAlign.TOP,textDirection:opt.dir,children:children.length?children:[new Paragraph('')]});}

function zipTable(paper,zip,boxes){
  const z=paper.zip,digits=String(zip||'').replace(/\D/g,'').split('');const widths=[],cells=[];
  z.x.forEach((x,i)=>{if(i>0){const gap=+(x-(z.x[i-1]+z.w)).toFixed(2);widths.push(gap);cells.push(cell(gap,[]));}
    widths.push(z.w);cells.push(cell(z.w,[new Paragraph({alignment:AlignmentType.CENTER,children:[run(digits[i]||'',paper.style==='zh'?16:14)]})],{vAlign:VerticalAlign.CENTER,borders:boxes?{top:BOX,bottom:BOX,left:BOX,right:BOX}:NO_BORDERS}));});
  return floatTable({x:z.x[0],y:z.top,widths,height:z.h,cells});
}

function jaPage(paper,rec,from,boxes){
  const {w,h}=paper,top=paper.zip.top+paper.zip.h+6,height=h-top-8;
  const addr=verticalNumbers(rec.address),name=withHonorific(rec.name,'ja');
  const addrW=w>110?30:24,nameW=w>110?30:24,sendW=w>110?24:18,side=(w-16-addrW-nameW-sendW)/2;
  const lines=s=>String(s||'').split(/\r\n|\r|\n/).filter(Boolean);
  const cells=[
    cell(sendW,lines(verticalNumbers(from)).map(s=>new Paragraph({alignment:AlignmentType.END,children:[run(s,9)]})),{dir:TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT}),
    cell(side,[]),
    cell(nameW,name?[new Paragraph({alignment:AlignmentType.CENTER,children:[run(name,w>110?28:24,{bold:true})]})]:[],{dir:TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT,vAlign:VerticalAlign.CENTER}),
    cell(side,[]),
    cell(addrW,addr?[new Paragraph({children:[run(addr,w>110?15:13)]})]:[],{dir:TextDirection.TOP_TO_BOTTOM_RIGHT_TO_LEFT,vAlign:VerticalAlign.BOTTOM})
  ];
  return [zipTable(paper,rec.zip,boxes),floatTable({x:8,y:top,widths:[sendW,side,nameW,side,addrW],height,cells}),new Paragraph('')];
}

function zhPage(paper,rec,from,boxes){
  const {w,h}=paper,x=20,width=w-40,name=withHonorific(rec.name,'zh');
  const tbl=new Table({layout:TableLayoutType.FIXED,width:{size:tw(width),type:WidthType.DXA},columnWidths:[tw(width)],borders:NO_BORDERS,
    float:{horizontalAnchor:TableAnchorType.PAGE,verticalAnchor:TableAnchorType.PAGE,absoluteHorizontalPosition:tw(x),absoluteVerticalPosition:tw(paper.zip.top+paper.zip.h+8),overlap:OverlapType.OVERLAP},
    rows:[[rec.address,16,{},AlignmentType.START,22],[name,24,{bold:true},AlignmentType.CENTER,26],[from,11,{},AlignmentType.END,h-paper.zip.top-paper.zip.h-8-48-10]].map(([text,size,opt,align,height])=>new TableRow({height:{value:tw(height),rule:HeightRule.EXACT},children:[cell(width,String(text||'').split(/\r\n|\r|\n/).filter(Boolean).map(s=>new Paragraph({alignment:align,children:[run(s,size,opt)]})),{vAlign:align===AlignmentType.END?VerticalAlign.BOTTOM:VerticalAlign.CENTER})]}))});
  return [zipTable(paper,rec.zip,boxes),tbl,new Paragraph('')];
}

// 返回 docx 的 sections（每个收件人一页）。
export function addressSections({paper:key,addresses,from='',zipBoxes=false,font}){
  const paper=PAPERS[key];if(!paper)throw Error('纸张无效。');
  const recs=String(addresses||'').split(/\r\n|\r|\n/).map(s=>parseAddress(s,paper.style)).filter(Boolean).slice(0,120);
  if(!recs.length)throw Error('请填写地址，每行一条。');
  // 横向纸张（中国 5 号信封）要标成 landscape：docx 按纵向尺寸传入再翻转，打印驱动才会选信封横放。
  const page={size:paper.w>paper.h?{width:tw(paper.h),height:tw(paper.w),orientation:PageOrientation.LANDSCAPE}:{width:tw(paper.w),height:tw(paper.h)},margin:{top:tw(5),bottom:tw(5),left:tw(5),right:tw(5)}};
  return recs.map(r=>({properties:{page},children:(paper.style==='zh'?zhPage:jaPage)(paper,r,from,zipBoxes)}));
}

// 地址标签的一格：邮编一行、地址、姓名加敬称（加粗大一号）。
export function labelParagraphs(line,style){
  const r=parseAddress(line,style);if(!r)return [new Paragraph('')];const out=[];
  if(r.zip)out.push(new Paragraph({children:[run(style==='zh'?'邮编 '+r.zip:'〒'+r.zip.slice(0,3)+'-'+r.zip.slice(3),10)]}));
  if(r.address)out.push(new Paragraph({children:[run(r.address,11)]}));
  if(r.name)out.push(new Paragraph({spacing:{before:80},children:[run(withHonorific(r.name,style),14,{bold:true})]}));
  return out.length?out:[new Paragraph('')];
}

// ── 寄件信息单／快递地址贴（配送ラベル）──────────────────────────────
// 快递公司各家面单格式不同，这里做通用版：A4 一张 4 份（虚线裁切），或 100×150 mm 热敏标签一张 1 份。
// 每行一条：邮编、地址、姓名、电话（空格或 Tab 分隔，顺序不限）。
const PHONE_RE=/(?<![\d-])(?:\+?\d{1,3}[- ]?)?(?:1[3-9]\d{9}|0\d{1,4}-\d{1,4}-\d{3,4}|0\d{9,10}|\d{3,4}-\d{7,8})(?![\d-])/;
export function parseShipping(line,style='ja'){
  const raw=toHalf(String(line||'').trim()).replace(/(?:电话|電話|TEL|Tel|tel)[:：]?\s*/g,' ').replace(/(?<!\d)(1[3-9]\d)[ -](\d{4})[ -](\d{4})(?!\d)/g,'$1$2$3');if(!raw)return null;
  let tel='',rest=raw;
  if(raw.includes('\t')){const f=raw.split('\t');const i=f.findIndex(x=>{const m=x.trim().match(PHONE_RE);return m&&m[0]===x.trim();});if(i>=0){tel=f[i].trim();f.splice(i,1);rest=f.join('\t');}}
  if(!tel){const m=rest.match(PHONE_RE);if(m){tel=m[0];rest=rest.replace(m[0],' ');}}
  const r=parseAddress(rest,style);return r?{...r,tel}:(tel?{zip:'',address:'',name:'',tel}:null);
}
const SHIP_PAPERS={a4:{w:210,h:297,cols:2,rows:2,margin:8},label100:{w:100,h:150,cols:1,rows:1,margin:5}};
const DASH={style:BorderStyle.DASHED,size:4,color:'9AA3AD'};
function shipBlock(rec,{from,note,style}){
  const L=style==='zh'?{to:'收件人',from:'寄件人',note:'物品／备注',tel:'电话',zip:'邮编 '}:{to:'お届け先',from:'ご依頼主',note:'品名・備考',tel:'TEL',zip:'〒'};
  const small=t=>new Paragraph({spacing:{after:40},children:[run(t,8,{color:'5D6878'})]});
  const out=[small(L.to)];
  if(rec.zip)out.push(new Paragraph({children:[run(style==='zh'?L.zip+rec.zip:L.zip+rec.zip.slice(0,3)+'-'+rec.zip.slice(3),11)]}));
  if(rec.address)out.push(new Paragraph({spacing:{after:80},children:[run(rec.address,12)]}));
  if(rec.name)out.push(new Paragraph({spacing:{after:60},children:[run(style==='zh'?rec.name:withHonorific(rec.name,'ja'),18,{bold:true})]}));
  if(rec.tel)out.push(new Paragraph({spacing:{after:160},children:[run(L.tel+'：'+rec.tel,12)]}));
  const f=String(from||'').split(/\r\n|\r|\n/).map(s=>s.trim()).filter(Boolean);
  if(f.length){out.push(new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:4,color:'C9CED6',space:4}},spacing:{before:120,after:40},children:[run(L.from,8,{color:'5D6878'})]}));f.forEach(s=>out.push(new Paragraph({children:[run(s,10)]})));}
  if(String(note||'').trim()){out.push(new Paragraph({spacing:{before:120,after:40},children:[run(L.note,8,{color:'5D6878'})]}));out.push(new Paragraph({children:[run(String(note).trim(),11)]}));}
  return out;
}
export function shippingSections({paper:key='a4',addresses,from='',note='',lang='zh'}){
  const p=SHIP_PAPERS[key]||SHIP_PAPERS.a4,style=lang==='ja'?'ja':'zh';
  const recs=String(addresses||'').split(/\r\n|\r|\n/).map(s=>parseShipping(s,style)).filter(Boolean).slice(0,120);
  if(!recs.length)throw Error('请填写地址，每行一条。');
  const page={size:{width:tw(p.w),height:tw(p.h)},margin:{top:tw(p.margin),bottom:tw(p.margin),left:tw(p.margin),right:tw(p.margin)}};
  const per=p.cols*p.rows,cw=(p.w-2*p.margin)/p.cols,ch=(p.h-2*p.margin-6)/p.rows,sections=[];
  for(let i=0;i<recs.length;i+=per){const group=recs.slice(i,i+per),rows=[];
    for(let r=0;r<p.rows;r++){const cells=[];for(let c=0;c<p.cols;c++){const rec=group[r*p.cols+c];
      cells.push(new TableCell({width:{size:tw(cw),type:WidthType.DXA},margins:{top:tw(5),bottom:tw(3),left:tw(5),right:tw(5)},borders:per>1?{top:DASH,bottom:DASH,left:DASH,right:DASH}:NO_BORDERS,children:rec?shipBlock(rec,{from,note,style}):[new Paragraph('')]}));}
      rows.push(new TableRow({height:{value:tw(ch),rule:HeightRule.EXACT},cantSplit:true,children:cells}));}
    sections.push({properties:{page},children:[new Table({layout:TableLayoutType.FIXED,width:{size:tw(cw*p.cols),type:WidthType.DXA},columnWidths:Array(p.cols).fill(tw(cw)),borders:NO_BORDERS,rows}),new Paragraph({spacing:{before:0,after:0,line:20,lineRule:'exact'},children:[new TextRun({text:'',size:2})]})]});}
  return sections;
}
