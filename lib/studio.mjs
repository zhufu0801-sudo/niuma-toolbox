import {pptThemes} from './ppt-themes.mjs';
// PPT 版式工作室：把大纲排成整套演示（封面、目录、章节页、多种内容版式、结束页），带切换和进入动画。
// 六套模板不只是换颜色：封面构图、章节页、卡片样式、编号方式、字体和动画都不同。
// 设计参考（只借鉴原则，未复制任何模板文件）：anthropics/skills 的 pptx 技能（配色与版式原则）、
// zarazhangrui/frontend-slides（风格预设思路）、hugohe3/ppt-master（原生可编辑形状）、willem4130 elite-powerpoint-designer（动画克制原则）。
import pptxgen from 'pptxgenjs';
import { execFile, spawn } from 'node:child_process';
import { promisify } from 'node:util';
import fs from 'node:fs/promises';
import path from 'node:path';
import { addMotion } from './pptx-motion.mjs';
const execFileAsync=promisify(execFile);
const W=1280,H=720;
const FONTS={zh:{sans:'Microsoft YaHei',serif:'SimSun'},ja:{sans:'Yu Gothic',serif:'Yu Mincho'}};
const WORDS={zh:{agenda:'目录',thanks:'谢谢观看',part:'PART',cont:'（续）'},ja:{agenda:'目次',thanks:'ご清聴ありがとうございました',part:'PART',cont:'（続き）'}};
const KANJI=['一','二','三','四','五','六','七','八','九','十','十一','十二','十三','十四','十五','十六','十七','十八','十九','二十'];
export const themes={
  midnight:{id:'midnight',name:'深海商务',nameJa:'ミッドナイト',desc:'稳重 · 汇报',c:{dark:'1E2761',paper:'FFFFFF',soft:'EEF1FA',ink:'1E2761',text:'2B3354',muted:'6B7394',accent:'F2B134',accent2:'CADCFC',onDark:'FFFFFF'},title:'sans',card:'panel',badge:'circle',transition:{type:'fade',speed:'slow'},motion:{title:'fade',item:'rise'}},
  neon:{id:'neon',name:'霓虹科技',nameJa:'ネオンテック',desc:'科技 · 发布会',c:{dark:'0B0F1A',paper:'0B0F1A',soft:'131B2D',ink:'F1F5FB',text:'C9D3E3',muted:'7D8AA3',accent:'22D3EE',accent2:'F472B6',onDark:'F1F5FB'},title:'sans',card:'outline',badge:'mono',transition:{type:'push',dir:'u',speed:'med'},motion:{title:'wipe',item:'wipe'}},
  swiss:{id:'swiss',name:'瑞士极简',nameJa:'スイスモダン',desc:'理性 · 答辩',c:{dark:'111111',paper:'FFFFFF',soft:'F2F2F2',ink:'111111',text:'2A2A2A',muted:'8D8D8D',accent:'E63946',accent2:'111111',onDark:'FFFFFF'},title:'sans',card:'rule',badge:'big',transition:{type:'wipe',dir:'r',speed:'med'},motion:{title:'wipe',item:'fade'}},
  forest:{id:'forest',name:'森林杂志',nameJa:'フォレスト',desc:'人文 · 品牌',c:{dark:'2C5F2D',paper:'FFFFFF',soft:'F1F6EC',ink:'2C5F2D',text:'34433A',muted:'6F7F72',accent:'97BC62',accent2:'D9E8C4',onDark:'FFFFFF'},title:'serif',card:'panel',badge:'serif',transition:{type:'fade',speed:'slow'},motion:{title:'fade',item:'fade'}},
  wa:{id:'wa',name:'和风雅致',nameJa:'和モダン',desc:'留白 · 日式',c:{dark:'2B2B2B',paper:'FFFFFF',soft:'F4F2EE',ink:'2B2B2B',text:'3A3A3A',muted:'9E9A93',accent:'C8442F',accent2:'264E70',onDark:'FFFFFF'},title:'serif',card:'rule',badge:'kanji',transition:{type:'fade',speed:'slow'},motion:{title:'fade',item:'rise'}},
  coral:{id:'coral',name:'活力珊瑚',nameJa:'コーラル',desc:'活泼 · 培训',c:{dark:'2F3C7E',paper:'FFFFFF',soft:'FFF1F1',ink:'2F3C7E',text:'2F3552',muted:'7A80A3',accent:'F96167',accent2:'F9E795',onDark:'FFFFFF'},title:'sans',card:'round',badge:'circle',transition:{type:'cover',dir:'l',speed:'med'},motion:{title:'rise',item:'zoom'}}
};
for(const scene of pptThemes){const base=themes[scene.decor];themes[scene.id]={...base,id:scene.id,decor:scene.decor,name:scene.zh,nameJa:scene.ja,c:{...base.c,dark:scene.bg,paper:'F3F2E9',soft:'E4E8E4',ink:scene.bg,text:scene.bg,muted:'617078',accent:scene.accent,accent2:scene.secondary,onDark:'F3F2E9'}};}
// 旧版的三个风格名继续可用
const ALIAS={business:'midnight',academic:'swiss',speech:'coral'};
const LAYOUTS={'要点':'bullets','列表':'bullets','箇条書き':'bullets',bullets:'bullets','卡片':'grid','カード':'grid',grid:'grid','数据':'stat','数字':'stat','数値':'stat',stat:'stat','时间线':'timeline','时间轴':'timeline','タイムライン':'timeline',timeline:'timeline','引言':'quote','金句':'quote','引用':'quote',quote:'quote','对比':'compare','比较':'compare','比較':'compare',compare:'compare','图表':'chart','グラフ':'chart',chart:'chart','图文':'image','画像':'image',image:'image'};
function clean(value){return String(value??'').replace(/\r/g,'').trim();}
// ---------- 图片与企业品牌 ----------
// 只接受 PNG / JPEG 的 data URL，每张不超过 2MB；读取像素尺寸用于等比缩放。
export function imageInfo(value,label='图片'){
  const m=/^data:image\/(png|jpe?g);base64,([A-Za-z0-9+/=]+)$/.exec(String(value||''));
  if(!m)throw Error(label+'只支持 PNG 或 JPG。');
  const buf=Buffer.from(m[2],'base64');if(buf.length>2*1024*1024)throw Error(label+'超过 2MB，请压缩后再用。');
  let width=0,height=0;
  if(buf.length>24&&buf.readUInt32BE(0)===0x89504e47){width=buf.readUInt32BE(16);height=buf.readUInt32BE(20);}
  else if(buf[0]===0xff&&buf[1]===0xd8){let i=2;while(i+9<buf.length){if(buf[i]!==0xff){i++;continue;}const mk=buf[i+1],len=buf.readUInt16BE(i+2);if(mk>=0xc0&&mk<=0xcf&&![0xc4,0xc8,0xcc].includes(mk)){height=buf.readUInt16BE(i+5);width=buf.readUInt16BE(i+7);break;}i+=2+len;}}
  if(!width||!height||width>20000||height>20000)throw Error(label+'无法读取尺寸，请换一张图片。');
  return {data:String(value),width,height};
}
function hex(c){return /^#?[0-9a-fA-F]{6}$/.test(String(c||''))?String(c).replace('#','').toUpperCase():'';}
function mix(a,b,t){const x=parseInt(a,16),y=parseInt(b,16),ch=s=>[s>>16&255,s>>8&255,s&255],p=ch(x),q=ch(y);return p.map((v,i)=>Math.round(v+(q[i]-v)*t).toString(16).padStart(2,'0')).join('').toUpperCase();}
function lum(c){const v=parseInt(c,16);return (0.2126*(v>>16&255)+0.7152*(v>>8&255)+0.0722*(v&255))/255;}
function normalizeBrand(b){
  if(!b||typeof b!=='object')return null;const color=hex(b.color),logo=b.logo?imageInfo(b.logo,'Logo'):null,name=clean(b.name).slice(0,40);
  if(b.color&&!color)throw Error('主色请用 #RRGGBB 格式，例如 #0B5FFF。');
  return color||logo?{color,logo,name}:null;
}
// 用企业主色替换模板的强调色；深色背景改为主色的深色版。颜色太浅时自动加深，保证白字看得清。
function brandTheme(t,brand){
  if(!brand||!brand.color)return t;
  const base=brand.color,accent=lum(base)>0.62?mix(base,'000000',0.35):base,deep=mix(accent,'000000',0.62),c={...t.c,accent};
  if(t.id==='midnight'||t.id==='forest'||t.id==='coral'){c.dark=deep;if(t.c.ink===t.c.dark)c.ink=deep;c.soft=mix(accent,'FFFFFF',0.9);}
  if(t.id==='midnight'||t.id==='forest')c.accent2=mix(accent,'FFFFFF',0.62);
  if(t.id==='swiss'||t.id==='wa')c.soft=mix(accent,'FFFFFF',0.92);
  return {...t,c};
}
export function normalizeStudio(input){
  if(!input||typeof input!=='object')throw Error('请填写制作内容。');
  const title=clean(input.title),subtitle=clean(input.subtitle),outline=clean(input.outline),raw=clean(input.theme||'midnight'),theme=ALIAS[raw]||raw;
  if(!title||title.length>80||/[\n\u0000-\u001f]/.test(title))throw Error('请填写 1～80 字的演示标题，不要换行。');
  if(subtitle.length>150||/[\u0000-\u001f]/.test(subtitle))throw Error('副标题最多 150 字，不要换行。');
  if(!Object.hasOwn(themes,theme))throw Error('请选择一种模板风格。');
  if(!outline||outline.length>20000||/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(outline))throw Error('请填写有效的大纲内容，最多 2 万字。');
  const lang=input.lang==='ja'||(input.lang!=='zh'&&/[぀-ヿ]/.test(title+outline))?'ja':'zh';
  const sections=outline.split(/\n[ \t]*\n+/).map(part=>{
    const lines=part.split('\n').map(line=>line.trim()).filter(Boolean);let heading=(lines.shift()||'').replace(/^#{1,6}\s+/,''),layout='';
    const tag=heading.match(/\s*[\[【]([^\]】]{1,12})[\]】]\s*$/);if(tag){layout=LAYOUTS[tag[1].trim()]||LAYOUTS[tag[1].trim().toLowerCase()]||'';heading=heading.slice(0,tag.index).trim();}
    const notes=lines.filter(line=>/^>/.test(line)).map(line=>line.replace(/^>\s*/,'')).join('\n').slice(0,1500);
    const points=lines.filter(line=>!/^>/.test(line)).map(line=>line.replace(/^[-*•・]\s+/,''));
    if(!heading||heading.length>100)throw Error('每节标题最多 100 字，请缩短标题。');
    if(!points.length)throw Error(`“${heading}”缺少内容，请在标题下方填写至少一行要点。`);
    if(points.some(p=>p.length>500))throw Error('单条要点超过 500 字，请拆成多行。');
    return {heading,points,layout,notes};
  });
  if(sections.length>20)throw Error('一次最多制作 20 个章节，请拆分大纲。');
  const images=(Array.isArray(input.images)?input.images:[]).slice(0,8).map((x,i)=>imageInfo(x,'第 '+(i+1)+' 张图片'));
  const brand=normalizeBrand(input.brand);
  return {title,subtitle,outline,theme,lang,sections,closing:input.closing!==false,images,brand};
}
export function wrapText(text,width,fontSize){
  const lines=[];let current='',weight=0;
  for(const char of Array.from(text)){
    if(char==='\n'){lines.push(current);current='';weight=0;continue;}
    const unit=char.codePointAt(0)>255?1:.58;
    if(weight+unit>width/fontSize&&current){lines.push(current);current='';weight=0;}
    current+=char;weight+=unit;
  }
  if(current)lines.push(current);return lines.length?lines:[''];
}
// ---------- 基本图形 ----------
const LH=1.3;
function newPage(role,background,t){return {role,background,elements:[],transition:t.transition};}
function add(p,el){el.name='e'+(p.elements.length+1);p.elements.push(el);return el;}
function rect(p,x,y,w,h,fill,o={}){return add(p,{type:'rect',x,y,w,h,fill,radius:o.radius||0,alpha:o.alpha||0,line:o.line||'',lineW:o.lineW||0,lineAlpha:o.lineAlpha||0,anim:o.anim||null});}
function ellipse(p,x,y,w,h,fill,o={}){return add(p,{type:'ellipse',x,y,w,h,fill,alpha:o.alpha||0,line:o.line||'',lineW:o.lineW||0,lineAlpha:o.lineAlpha||0,anim:o.anim||null});}
function line(p,x,y,w,h,color,o={}){return add(p,{type:'line',x,y,w,h,color,width:o.width||1,alpha:o.alpha||0,anim:o.anim||null});}
function text(p,ctx,value,x,y,w,h,o={}){
  let size=o.size||26;const min=o.min||size,lh=o.lh||LH;let lines=wrapText(value,w,size);
  while(lines.length*size*lh>h&&size>min){size-=2;lines=wrapText(value,w,size);}
  if(lines.length*size*lh>h)throw Error('某段文字无法完整排入页面，请缩短标题或拆分要点。');
  const face=FONTS[ctx.lang][o.font==='title'?(ctx.t.title):(o.font||'sans')]||FONTS[ctx.lang].sans;
  return add(p,{type:'text',text:value,lines,x,y,w,h,fontSize:size,color:o.color||ctx.t.c.text,bold:!!o.bold,italic:!!o.italic,align:o.align||'left',face,lh,alpha:o.alpha||0,anim:o.anim||null});
}
const anim=(effect,delay=0,dur=600)=>({effect,delay,dur});
function split(point){const m=point.match(/^(.{1,14}?)[：:]\s*(.+)$/);return m?{label:m[1].trim(),desc:m[2].trim()}:{label:'',desc:point};}
const NUM=/^([+\-−]?\d[\d,.]*\s*(?:%|％|万|亿|倍|x|X|k|K|M|B|\+)?)\s*[：:，,、\-—]?\s*(.*)$/;
const TIME=/^((?:(?:19|20)\d{2}(?:[年.\/-]\d{1,2}(?:月)?)?|\d{1,2}月(?:\d{1,2}日)?|Q[1-4]|H[12]|第[一二三四五六七八九十\d]+[阶段步周月年期季]|Phase\s?\d+|Step\s?\d+|STEP\s?\d+)[年]?)\s*[：:，,、\-—]?\s*(.*)$/;
function badgeText(t,i,lang){const n=String(i+1).padStart(2,'0');return t.badge==='kanji'?KANJI[i]||String(i+1):t.badge==='mono'?'['+n+']':n;}
// ---------- 每套模板自己的封面、章节页、内容页页眉、结束页 ----------
const DECOR={
  midnight:{
    cover(p,ctx,r){const c=ctx.t.c;ellipse(p,830,110,400,400,c.accent2,{alpha:78});ellipse(p,960,300,300,300,c.accent2,{alpha:60});ellipse(p,880,420,140,140,c.accent,{alpha:10});
      rect(p,88,176,56,6,c.accent);text(p,ctx,r.title,88,206,700,240,{size:62,min:34,bold:true,color:c.onDark,font:'title',anim:anim('fade',0,800)});
      if(r.subtitle)text(p,ctx,r.subtitle,90,470,700,96,{size:26,min:18,color:c.accent2,anim:anim('rise',300)});},
    divider(p,ctx,i,h){const c=ctx.t.c;text(p,ctx,String(i+1).padStart(2,'0'),88,120,520,280,{size:220,color:c.accent2,alpha:55,bold:true,lh:1.1});rect(p,96,430,56,6,c.accent);text(p,ctx,h,96,460,1000,150,{size:54,min:32,bold:true,color:c.onDark,font:'title',anim:anim('rise',0,700)});},
    header(p,ctx,h,i){const c=ctx.t.c;if(i>=0)text(p,ctx,String(i+1).padStart(2,'0'),72,38,200,30,{size:18,bold:true,color:c.accent});return text(p,ctx,h,72,66,1136,84,{size:42,min:28,bold:true,color:c.ink,font:'title',anim:anim('fade')});},
    closing(p,ctx,r){const c=ctx.t.c;ellipse(p,860,140,360,360,c.accent2,{alpha:80});text(p,ctx,WORDS[ctx.lang].thanks,88,250,900,120,{size:64,min:36,bold:true,color:c.onDark,font:'title',anim:anim('fade',0,800)});text(p,ctx,r.title,92,390,900,60,{size:24,min:16,color:c.accent2,anim:anim('rise',300)});}
  },
  neon:{
    grid(p,ctx){const c=ctx.t.c;for(let x=160;x<W;x+=160)line(p,x,0,0,H,c.accent,{alpha:90,width:.75});for(let y=120;y<H;y+=120)line(p,0,y,W,0,c.accent,{alpha:92,width:.75});},
    cover(p,ctx,r){const c=ctx.t.c;this.grid(p,ctx);rect(p,0,560,W,3,c.accent);rect(p,0,566,420,3,c.accent2);ellipse(p,1040,90,150,150,c.accent2,{alpha:70});ellipse(p,1080,130,70,70,c.accent,{alpha:20});
      text(p,ctx,'> '+(r.subtitle?'LIVE':'READY'),88,140,300,34,{size:20,bold:true,color:c.accent});text(p,ctx,r.title,88,190,1000,250,{size:66,min:34,bold:true,color:c.ink,font:'title',anim:anim('wipe',0,900)});
      if(r.subtitle)text(p,ctx,r.subtitle,88,600,1000,70,{size:24,min:16,color:c.text,anim:anim('fade',400)});},
    divider(p,ctx,i,h){const c=ctx.t.c;this.grid(p,ctx);text(p,ctx,'['+String(i+1).padStart(2,'0')+']',88,150,700,230,{size:180,bold:true,color:c.accent,lh:1.1,anim:anim('wipe',0,700)});rect(p,96,430,220,4,c.accent2);text(p,ctx,h,96,460,1080,150,{size:52,min:30,bold:true,color:c.ink,font:'title',anim:anim('fade',300)});},
    header(p,ctx,h,i){const c=ctx.t.c;this.grid(p,ctx);if(i>=0)text(p,ctx,'// '+String(i+1).padStart(2,'0'),72,36,240,30,{size:18,bold:true,color:c.accent2});return text(p,ctx,h,72,64,1100,84,{size:42,min:28,bold:true,color:c.ink,font:'title',anim:anim('wipe')});},
    closing(p,ctx,r){const c=ctx.t.c;this.grid(p,ctx);text(p,ctx,WORDS[ctx.lang].thanks,88,250,1100,120,{size:64,min:34,bold:true,color:c.accent,font:'title',anim:anim('wipe',0,900)});rect(p,92,392,300,3,c.accent2);text(p,ctx,r.title,92,410,1000,60,{size:24,min:16,color:c.text,anim:anim('fade',400)});}
  },
  swiss:{
    cover(p,ctx,r){const c=ctx.t.c;rect(p,0,0,460,H,c.accent);text(p,ctx,ctx.lang==='ja'?'セクション':'章节',44,392,380,40,{size:22,color:'FFFFFF'});text(p,ctx,String(r.sections.length).padStart(2,'0'),40,430,400,250,{size:220,bold:true,color:'FFFFFF',lh:1.05});
      text(p,ctx,r.title,540,170,660,300,{size:60,min:32,bold:true,color:c.ink,font:'title',anim:anim('wipe',0,800)});rect(p,540,500,660,2,c.ink);if(r.subtitle)text(p,ctx,r.subtitle,540,520,660,90,{size:24,min:16,color:c.muted,anim:anim('fade',300)});},
    divider(p,ctx,i,h){const c=ctx.t.c;p.background=c.accent;text(p,ctx,String(i+1).padStart(2,'0'),72,60,700,320,{size:260,bold:true,color:'FFFFFF',lh:1.05});rect(p,80,440,1120,3,'FFFFFF');text(p,ctx,h,80,470,1120,160,{size:56,min:32,bold:true,color:'FFFFFF',font:'title',anim:anim('wipe',0,700)});},
    header(p,ctx,h,i){const c=ctx.t.c;rect(p,1184,56,24,24,c.accent);const t=text(p,ctx,h,72,48,1080,96,{size:46,min:28,bold:true,color:c.ink,font:'title',anim:anim('wipe')});rect(p,72,152,1136,2,c.ink);return t;},
    closing(p,ctx,r){const c=ctx.t.c;rect(p,0,0,W,300,c.accent);text(p,ctx,WORDS[ctx.lang].thanks,72,140,1136,120,{size:66,min:34,bold:true,color:'FFFFFF',font:'title',anim:anim('wipe',0,800)});text(p,ctx,r.title,72,360,1136,60,{size:26,min:16,bold:true,color:c.ink,anim:anim('fade',300)});}
  },
  forest:{
    cover(p,ctx,r){const c=ctx.t.c;ellipse(p,860,380,400,300,c.accent,{alpha:70});ellipse(p,1010,300,240,240,c.accent2,{alpha:80});rect(p,88,150,3,380,c.accent);
      text(p,ctx,r.title,120,180,760,250,{size:60,min:32,bold:true,color:c.onDark,font:'title',anim:anim('fade',0,900)});if(r.subtitle)text(p,ctx,r.subtitle,122,450,760,90,{size:24,min:16,italic:true,color:c.accent2,font:'serif',anim:anim('fade',400)});},
    divider(p,ctx,i,h){const c=ctx.t.c;p.background=c.soft;text(p,ctx,'Chapter '+String(i+1).padStart(2,'0'),96,170,600,60,{size:30,italic:true,color:c.muted,font:'serif'});rect(p,96,250,80,3,c.accent);text(p,ctx,h,96,280,1000,220,{size:58,min:32,bold:true,color:c.ink,font:'title',anim:anim('fade',0,800)});},
    header(p,ctx,h,i){const c=ctx.t.c;if(i>=0)text(p,ctx,'Chapter '+String(i+1).padStart(2,'0'),72,34,300,32,{size:18,italic:true,color:c.muted,font:'serif'});return text(p,ctx,h,72,66,1136,84,{size:42,min:28,bold:true,color:c.ink,font:'title',anim:anim('fade')});},
    closing(p,ctx,r){const c=ctx.t.c;ellipse(p,920,360,300,300,c.accent,{alpha:70});rect(p,88,200,3,240,c.accent);text(p,ctx,WORDS[ctx.lang].thanks,120,230,900,120,{size:60,min:34,bold:true,color:c.onDark,font:'title',anim:anim('fade',0,900)});text(p,ctx,r.title,122,370,900,60,{size:24,min:16,italic:true,color:c.accent2,font:'serif',anim:anim('fade',400)});}
  },
  wa:{
    seal(p,ctx,x,y,s,ch){const c=ctx.t.c;rect(p,x,y,s,s,c.accent);text(p,ctx,ch,x,y+s*.14,s,s*.8,{size:Math.round(s*.56),bold:true,color:'FFFFFF',font:'serif',align:'center',lh:1.2});},
    cover(p,ctx,r){const c=ctx.t.c;this.seal(p,ctx,96,96,92,Array.from(r.title)[0]);line(p,1160,90,0,540,c.muted,{width:1});line(p,1176,150,0,420,c.accent,{width:2});
      text(p,ctx,r.title,96,270,980,230,{size:58,min:32,bold:true,color:c.ink,font:'title',anim:anim('fade',0,1000)});if(r.subtitle)text(p,ctx,r.subtitle,98,520,980,80,{size:24,min:16,color:c.muted,anim:anim('rise',500)});},
    divider(p,ctx,i,h){const c=ctx.t.c;text(p,ctx,KANJI[i]||String(i+1),96,110,500,300,{size:220,color:c.accent,font:'serif',lh:1.1,anim:anim('fade',0,900)});line(p,100,450,0,170,c.accent,{width:3});text(p,ctx,h,130,460,1000,160,{size:52,min:30,bold:true,color:c.ink,font:'title',anim:anim('rise',300)});},
    header(p,ctx,h,i){const c=ctx.t.c;rect(p,72,58,6,72,c.accent);if(i>=0)text(p,ctx,KANJI[i]||String(i+1),1120,48,90,60,{size:40,color:c.muted,font:'serif',align:'right'});return text(p,ctx,h,100,56,1000,84,{size:42,min:28,bold:true,color:c.ink,font:'title',anim:anim('fade')});},
    closing(p,ctx,r){const c=ctx.t.c;this.seal(p,ctx,594,170,92,'謝');text(p,ctx,WORDS[ctx.lang].thanks,140,320,1000,100,{size:52,min:30,bold:true,color:c.ink,font:'title',align:'center',anim:anim('fade',0,1000)});text(p,ctx,r.title,140,430,1000,60,{size:22,min:16,color:c.muted,align:'center',anim:anim('fade',500)});}
  },
  coral:{
    cover(p,ctx,r){const c=ctx.t.c;ellipse(p,800,150,420,420,c.accent);ellipse(p,1060,80,160,160,c.accent2);ellipse(p,760,470,120,120,c.accent2,{alpha:40});
      text(p,ctx,r.title,88,200,680,260,{size:60,min:32,bold:true,color:c.onDark,font:'title',anim:anim('rise',0,800)});if(r.subtitle)text(p,ctx,r.subtitle,90,480,660,90,{size:26,min:18,color:c.accent2,anim:anim('rise',300)});},
    divider(p,ctx,i,h){const c=ctx.t.c;p.background=c.accent;ellipse(p,90,110,240,240,'FFFFFF');text(p,ctx,String(i+1),90,160,240,150,{size:110,bold:true,color:c.accent,align:'center',lh:1.1,anim:anim('zoom',0,600)});text(p,ctx,h,96,420,1080,180,{size:56,min:32,bold:true,color:'FFFFFF',font:'title',anim:anim('rise',300)});},
    header(p,ctx,h,i){const c=ctx.t.c;ellipse(p,1140,30,90,90,c.accent2);ellipse(p,1190,80,40,40,c.accent,{alpha:20});return text(p,ctx,h,72,52,1040,96,{size:42,min:28,bold:true,color:c.ink,font:'title',anim:anim('rise')});},
    closing(p,ctx,r){const c=ctx.t.c;ellipse(p,880,120,320,320,c.accent);ellipse(p,1080,380,140,140,c.accent2);text(p,ctx,WORDS[ctx.lang].thanks,88,250,780,120,{size:62,min:34,bold:true,color:c.onDark,font:'title',anim:anim('rise',0,800)});text(p,ctx,r.title,92,390,760,60,{size:24,min:16,color:c.accent2,anim:anim('fade',400)});}
  }
};
// ---------- 内容版式 ----------
const BOX={x:72,y:184,w:1136,h:456};
function cardBg(p,ctx,x,y,w,h,i){const t=ctx.t,c=t.c;
  if(t.card==='panel')rect(p,x,y,w,h,c.soft);
  else if(t.card==='round')rect(p,x,y,w,h,i%2?c.soft:'FFF8DC',{radius:18});
  else if(t.card==='outline')rect(p,x,y,w,h,c.soft,{line:c.accent,lineW:1,lineAlpha:55});
  else rect(p,x,y,w,2,t.id==='wa'?c.muted:c.ink,{alpha:t.id==='wa'?40:0});}
function badge(p,ctx,x,y,i,size=44){const t=ctx.t,c=t.c,label=badgeText(t,i,ctx.lang);
  if(t.badge==='circle'){ellipse(p,x,y,size,size,c.accent);text(p,ctx,String(i+1),x,y+size*.18,size,size*.7,{size:Math.round(size*.46),bold:true,color:t.id==='midnight'?c.dark:'FFFFFF',align:'center',lh:1.2});return size+16;}
  const sz=t.badge==='big'?34:t.badge==='kanji'?30:t.badge==='serif'?30:22;text(p,ctx,label,x,y,label.length*sz*.8+20,sz*1.3,{size:sz,bold:t.badge!=='kanji',color:c.accent,font:t.badge==='kanji'||t.badge==='serif'?'serif':'sans',lh:1.3});return 0;}
function bullets(pages,ctx,sec,mk){
  let p=mk(false),y=BOX.y,n=0;
  sec.points.forEach((point)=>{
    const lines=wrapText(point,BOX.w-140,28),height=Math.max(80,Math.ceil(lines.length*28*LH)+34);
    if(height>BOX.h)throw Error('单条要点排版后过长，请拆为多行较短的要点。');
    if(y+height>BOX.y+BOX.h){p=mk(true);y=BOX.y;n=0;}
    cardBg(p,ctx,BOX.x,y,BOX.w,height,n);
    const off=ctx.t.card==='rule'?0:18;
    if(ctx.t.badge==='circle')badge(p,ctx,BOX.x+off+6,y+(height-40)/2,n,40);else badge(p,ctx,BOX.x+off+6,y+17,n);
    text(p,ctx,point,BOX.x+110,y+17,BOX.w-140,height-24,{size:28,color:ctx.t.c.text,anim:anim(ctx.t.motion.item,250+n*160)});
    y+=height+(ctx.t.card==='rule'?6:14);n++;
  });
}
function grid(pages,ctx,sec,mk){
  const n=sec.points.length,cols=n<=3?n:n===4?2:3,rows=Math.ceil(n/cols),gap=24,cw=(BOX.w-gap*(cols-1))/cols,maxH=(BOX.h-gap*(rows-1))/rows,p=mk(false),pad=26;
  const items=sec.points.map(point=>{const s=split(point),size=s.label?24:28,lines=wrapText(s.desc,cw-pad*2,size);return {...s,size,need:pad*2+60+(s.label?50:0)+Math.ceil(lines.length*size*LH)+8};});
  const rowH=[];for(let r=0;r<rows;r++)rowH.push(Math.min(maxH,Math.max(170,...items.slice(r*cols,r*cols+cols).map(x=>x.need))));
  let y=BOX.y+Math.max(0,(BOX.h-rowH.reduce((a,b)=>a+b,0)-gap*(rows-1))/3);
  for(let r=0;r<rows;r++){items.slice(r*cols,r*cols+cols).forEach((it,k)=>{const i=r*cols+k,x=BOX.x+k*(cw+gap),ch=rowH[r],delay=250+i*150,eff=ctx.t.motion.item;
      cardBg(p,ctx,x,y,cw,ch,i);let ty=y+pad+(ctx.t.card==='rule'?6:0);
      badge(p,ctx,x+pad,ty,i,40);ty+=60;
      if(it.label){text(p,ctx,it.label,x+pad,ty,cw-pad*2,44,{size:30,min:20,bold:true,color:ctx.t.c.ink,anim:anim(eff,delay)});ty+=50;}
      text(p,ctx,it.desc,x+pad,ty,cw-pad*2,y+ch-ty-10,{size:it.size,min:16,color:ctx.t.c.text,anim:anim(eff,delay+60)});});
    y+=rowH[r]+gap;}
}
function stat(pages,ctx,sec,mk){
  const n=sec.points.length,gap=28,cw=(BOX.w-gap*(n-1))/n,p=mk(false),c=ctx.t.c;
  sec.points.forEach((point,i)=>{const m=point.match(NUM),x=BOX.x+i*(cw+gap),y=BOX.y+40;
    cardBg(p,ctx,x,y,cw,340,i);
    text(p,ctx,m[1].replace(/\s+/g,''),x+24,y+50,cw-48,120,{size:n>3?64:80,min:36,bold:true,color:c.accent,lh:1.15,anim:anim(ctx.t.id==='coral'?'zoom':ctx.t.motion.item,250+i*200,700)});
    if(m[2])text(p,ctx,m[2],x+24,y+190,cw-48,130,{size:26,min:16,color:c.text,anim:anim('fade',400+i*200)});});
}
function timeline(pages,ctx,sec,mk){
  const n=sec.points.length,p=mk(false),c=ctx.t.c,y0=BOX.y+170,step=BOX.w/n;
  line(p,BOX.x,y0+14,BOX.w,0,ctx.t.id==='neon'?c.accent:c.muted,{width:2,anim:anim('wipe',150,900)});
  sec.points.forEach((point,i)=>{const m=point.match(TIME)||[point,'',point],x=BOX.x+i*step,cx=x+step/2,delay=400+i*220;
    ellipse(p,cx-14,y0,28,28,i===n-1?c.accent:(ctx.t.id==='swiss'?c.ink:c.accent),{alpha:i===n-1?0:25,anim:anim('zoom',delay,400)});
    if(m[1])text(p,ctx,m[1],x+8,y0-110,step-16,90,{size:30,min:18,bold:true,color:c.ink,align:'center',anim:anim('fade',delay)});
    text(p,ctx,m[2]||point,x+12,y0+56,step-24,BOX.y+BOX.h-y0-56,{size:24,min:15,color:c.text,align:'center',anim:anim('fade',delay+80)});});
}
function quote(pages,ctx,sec,mk){
  const p=mk(false),c=ctx.t.c,q=sec.points.join('\n'),parts=q.split(/\s*[—–-]{2}\s*/);
  text(p,ctx,'“',BOX.x,BOX.y-30,160,160,{size:160,bold:true,color:c.accent,font:'serif',lh:1,alpha:20});
  text(p,ctx,parts[0],BOX.x+60,BOX.y+70,BOX.w-120,280,{size:36,min:22,color:c.ink,font:'serif',lh:1.45,anim:anim('fade',200,900)});
  if(parts[1])text(p,ctx,'—— '+parts[1],BOX.x+60,BOX.y+370,BOX.w-120,60,{size:22,min:16,color:c.muted,align:'right',anim:anim('fade',700)});
}
const PAIR=/^(.{1,30}?)\s*[：:]\s*([+\-]?\d[\d,]*(?:\.\d+)?)\s*(%|％|万|亿|元|円|件|人|个|次|台|小时|時間|h|k|K|M)?$/;
const CMP=/^(.{1,30}?)\s*[｜|]\s*(.{1,30})$/;
function compare(pages,ctx,sec,mk){
  let pts=[...sec.points],titles;const m=pts[0].match(CMP);
  if(m&&pts.length>=3){titles=[m[1].trim(),m[2].trim()];pts=pts.slice(1);}else titles=ctx.lang==='ja'?['A 案','B 案']:['方案 A','方案 B'];
  const half=Math.ceil(pts.length/2),cols=[pts.slice(0,half),pts.slice(half)],gap=88,cw=(BOX.w-gap)/2,p=mk(false),c=ctx.t.c;
  // 卡片高度按内容决定（条目少时不留大片空白），整体在内容区垂直居中
  const sizes=cols.map(items=>items.map(it=>Math.ceil(wrapText(it,cw-90,24).length*24*LH)+14));
  const need=Math.max(...sizes.map(hs=>96+hs.reduce((a,h)=>a+h+8,0)+24));
  if(need>BOX.h)throw Error('对比内容太多，请减少条目。');
  const ch=Math.max(260,need),y0=BOX.y+Math.round((BOX.h-ch)/3);
  cols.forEach((items,k)=>{const x=BOX.x+k*(cw+gap),eff=ctx.t.motion.item;
    rect(p,x,y0,cw,ch,c.soft,{radius:ctx.t.card==='round'?18:0});rect(p,x,y0,cw,6,k?c.accent:c.muted);
    text(p,ctx,titles[k],x+28,y0+28,cw-56,46,{size:30,min:20,bold:true,color:k?c.accent:c.ink,anim:anim(eff,250+k*500)});
    let y=y0+96;items.forEach((it,i)=>{const h=sizes[k][i];
      ellipse(p,x+30,y+10,12,12,k?c.accent:c.muted);text(p,ctx,it,x+56,y,cw-90,h,{size:24,min:16,color:c.text,anim:anim(eff,350+k*500+i*120)});y+=h+8;});});
  const cx=BOX.x+cw+gap/2,cy=y0+ch/2;ellipse(p,cx-30,cy-30,60,60,c.accent,{anim:anim('zoom',200,500)});text(p,ctx,'VS',cx-30,cy-15,60,30,{size:22,bold:true,color:'FFFFFF',align:'center',lh:1.2});
}
function chart(pages,ctx,sec,mk){
  const rows=sec.points.map(x=>x.match(PAIR)),p=mk(false),c=ctx.t.c,unit=(rows[0][3]||'').replace('％','%');
  add(p,{type:'chart',x:BOX.x,y:BOX.y,w:BOX.w,h:BOX.h-8,labels:rows.map(m=>m[1].trim()),values:rows.map(m=>Number(m[2].replace(/,/g,''))),unit,color:c.accent,text:c.text,muted:c.muted,face:FONTS[ctx.lang].sans,anim:anim('fade',200,800)});
}
function imageLayout(pages,ctx,sec,mk){
  const img=ctx.images.shift(),p=mk(false),c=ctx.t.c,bw=560,bx=BOX.x,by=BOX.y,bh=BOX.h;
  if(img){const scale=Math.min(bw/img.width,bh/img.height),w=img.width*scale,h=img.height*scale;add(p,{type:'image',data:img.data,x:bx+(bw-w)/2,y:by+(bh-h)/2,w,h,anim:anim('fade',150,700)});}
  else{rect(p,bx,by,bw,bh,c.soft,{line:c.muted,lineW:1,lineAlpha:40});text(p,ctx,ctx.lang==='ja'?'ここに画像（制作ページで画像を追加すると自動で入ります）':'这里放图片（在制作页添加图片后会自动放入）',bx+40,by+bh/2-30,bw-80,70,{size:20,min:14,color:c.muted,align:'center'});}
  let y=by+10;const tx=bx+bw+48,tw=BOX.w-bw-48;
  sec.points.forEach((it,i)=>{const lines=wrapText(it,tw-28,24),h=Math.ceil(lines.length*24*LH)+16;if(y+h>by+bh)throw Error('图文页文字太多，请减少条目。');
    rect(p,tx,y+6,6,Math.min(28,h-16),c.accent);text(p,ctx,it,tx+24,y,tw-28,h,{size:24,min:16,color:c.text,anim:anim(ctx.t.motion.item,300+i*150)});y+=h+10;});
}
const RENDER={bullets,grid,stat,timeline,quote,compare,chart,image:imageLayout};
export function chooseLayout(sec){
  const pts=sec.points,n=pts.length;
  if(sec.layout)return sec.layout;
  if(n===1&&pts[0].length>=16)return 'quote';
  if(n>=3&&n<=6&&pts.filter(x=>TIME.test(x)).length>=n-1&&pts.every(x=>x.length<=60))return 'timeline';
  if(n>=3&&n<=8&&pts.every(x=>PAIR.test(x)))return 'chart';
  if(n>=3&&CMP.test(pts[0])&&pts.slice(1).every(x=>!CMP.test(x)))return 'compare';
  if(n>=2&&n<=4&&pts.every(x=>NUM.test(x)&&!TIME.test(x)&&x.length<=40))return 'stat';
  if(n>=3&&n<=6&&pts.every(x=>x.length<=48))return 'grid';
  return 'bullets';
}
function layoutFits(kind,sec){const n=sec.points.length;return kind==='bullets'||kind==='image'&&n<=6||kind==='compare'&&n>=2&&n<=13||(kind==='chart'&&n>=2&&n<=12&&sec.points.every(x=>PAIR.test(x)))||(kind==='grid'&&n>=2&&n<=6)||(kind==='stat'&&n>=1&&n<=4&&sec.points.every(x=>NUM.test(x)))||(kind==='timeline'&&n>=2&&n<=6)||(kind==='quote'&&n<=2);}
// 企业 Logo：封面和结束页放大一些，内容页放在页脚左侧；深色或彩色底上垫一块白底，保证 Logo 看得清。
function placeLogo(p,ctx,where){
  const L=ctx.brand&&ctx.brand.logo;if(!L)return 0;const id=ctx.t.id;
  let pos=where==='footer'?{x:72,y:664,h:30,max:150}:{x:88,y:52,h:52,max:240};
  if(where!=='footer'&&id==='swiss'&&where==='cover')pos={x:540,y:72,h:52,max:240};
  if(where==='cover'&&id==='wa')pos={x:98,y:612,h:46,max:220};
  const w=Math.min(pos.max,pos.h*L.width/L.height),h=w*L.height/L.width,plate=p.background!=='FFFFFF'||(id==='swiss'&&where==='closing'),pad=plate?8:0;
  if(plate)rect(p,pos.x-pad,pos.y-pad,w+pad*2,h+pad*2,'FFFFFF',{radius:8});
  add(p,{type:'image',data:L.data,x:pos.x,y:pos.y,w,h,logo:true});
  return w+pad*2+16;
}
export function planStudio(input){
  const r=normalizeStudio(input),t=brandTheme(themes[r.theme],r.brand),ctx={t,lang:r.lang,images:[...r.images],brand:r.brand},pages=[],dark=t.c.dark,words=WORDS[r.lang];
  const deco=DECOR[t.decor||t.id];
  const bgOf=k=>({swiss:{cover:t.c.paper,divider:t.c.accent,closing:t.c.paper},wa:{cover:t.c.paper,divider:t.c.paper,closing:t.c.paper}}[t.id]||{})[k]||dark;
  const cover=newPage('cover',bgOf('cover'),t);deco.cover(cover,ctx,r);pages.push(cover);
  const many=r.sections.length>=3;
  if(many){const p=newPage('agenda',t.c.paper,t);deco.header(p,ctx,words.agenda,-1);
    const cols=r.sections.length>5?2:1,per=Math.ceil(r.sections.length/cols),rowH=Math.min(84,BOX.h/per);
    r.sections.forEach((s,i)=>{const col=Math.floor(i/per),row=i%per,x=BOX.x+col*(BOX.w/cols),y=BOX.y+row*rowH,w=BOX.w/cols-24;
      text(p,ctx,t.badge==='kanji'?KANJI[i]:String(i+1).padStart(2,'0'),x,y+8,90,rowH-12,{size:Math.min(40,rowH*.5),min:18,bold:true,color:t.c.accent,font:t.badge==='kanji'||t.badge==='serif'?'serif':'sans',lh:1.2,anim:anim('fade',200+i*120)});
      text(p,ctx,s.heading,x+100,y+12,w-100,rowH-16,{size:Math.min(28,rowH*.36),min:16,color:t.c.ink,bold:true,lh:1.25,anim:anim(t.motion.item,260+i*120)});
      if(t.card==='rule'||t.card==='outline')line(p,x,y+rowH-4,w,0,t.c.muted,{alpha:60});});
    pages.push(p);}
  r.sections.forEach((sec,i)=>{
    if(many){const d=newPage('divider',bgOf('divider'),t);deco.divider(d,ctx,i,sec.heading);pages.push(d);}
    const mk=(cont)=>{const p=newPage('content',t.c.paper,t);deco.header(p,ctx,sec.heading+(cont?words.cont:''),i);pages.push(p);return p;};
    let kind=chooseLayout(sec);if(!layoutFits(kind,sec))kind='bullets';
    const mark=pages.length;
    try{RENDER[kind](pages,ctx,sec,mk);}catch(e){if(kind==='bullets')throw e;pages.length=mark;bullets(pages,ctx,sec,mk);kind='bullets';}
    for(let k=mark;k<pages.length;k++)pages[k].layout=kind;
    if(sec.notes&&pages[mark])pages[mark].notes=sec.notes;
  });
  if(r.closing){const p=newPage('closing',bgOf('closing'),t);deco.closing(p,ctx,r);pages.push(p);}
  if(pages.length>40)throw Error('自动分页后超过 40 页，请减少内容或分批制作。');
  pages.forEach((p,index)=>{if(p.role==='cover'||p.role==='closing'){placeLogo(p,ctx,p.role);return;}if(p.role!=='content'&&p.role!=='agenda')return;const color=t.c.muted,lw=placeLogo(p,ctx,'footer');
    text(p,ctx,`${String(index+1).padStart(2,'0')} / ${String(pages.length).padStart(2,'0')}`,1096,674,112,26,{size:15,color,align:'right'});
    text(p,ctx,r.title,72+lw,674,900-lw,26,{size:14,min:10,color});});
  for(const p of pages)for(const e of p.elements){if(e.x<0||e.y<0||e.x+e.w>W+0.5||e.y+e.h>H+0.5)throw Error('版式计算超出页面，请缩短内容。');}
  return {title:r.title,theme:t.id,lang:r.lang,width:W,height:H,pages,
    warnings:r.lang==='ja'?['ローカルのテンプレートでレイアウトします。内容の追加や事実確認は行いません。','切り替え効果とアニメーションは PowerPoint / WPS で再生されます（プレビューでは静止表示）。']:['这是本地模板排版，不会自动扩写或核实内容。','切换效果和进入动画在 PowerPoint / WPS 放映时播放，预览为静态画面。'],local:true};
}
export async function exportStudio(input){
  const plan=planStudio(input),ppt=new pptxgen();ppt.defineLayout({name:'STUDIO',width:13.333333,height:7.5});ppt.layout='STUDIO';ppt.title=plan.title;ppt.subject='本地大纲与模板制作';ppt.author='办公工具箱';ppt.lang=plan.lang==='ja'?'ja-JP':'zh-CN';
  const none=fill=>({color:fill,transparency:100});
  for(const page of plan.pages){const slide=ppt.addSlide();slide.background={color:page.background};if(page.notes)slide.addNotes(page.notes);
    for(const e of page.elements){const g={x:e.x/96,y:e.y/96,w:e.w/96,h:e.h/96,objectName:e.name};
      if(e.type==='rect'||e.type==='ellipse'){const shape=e.type==='ellipse'?ppt.ShapeType.ellipse:e.radius?ppt.ShapeType.roundRect:ppt.ShapeType.rect;
        slide.addShape(shape,{...g,...(e.radius&&e.type==='rect'?{rectRadius:e.radius/96}:{}),fill:{color:e.fill,transparency:e.alpha||0},line:e.line?{color:e.line,width:e.lineW*.75,transparency:e.lineAlpha||0}:none(e.fill)});}
      else if(e.type==='line')slide.addShape(ppt.ShapeType.line,{...g,line:{color:e.color,width:e.width*.75,transparency:e.alpha||0}});
      else if(e.type==='image')slide.addImage({...g,data:e.data.replace(/^data:/,'')});
      else if(e.type==='chart')slide.addChart(ppt.ChartType.bar,[{name:plan.title,labels:e.labels,values:e.values}],{...g,barDir:'col',chartColors:[e.color],showValue:true,dataLabelColor:e.text,dataLabelFontSize:12,dataLabelFontFace:e.face,dataLabelPosition:'outEnd',catAxisLabelColor:e.text,catAxisLabelFontFace:e.face,catAxisLabelFontSize:14,valAxisLabelColor:e.muted,valAxisLabelFontSize:11,valGridLine:{color:e.muted,size:0.5,style:'dash'},catAxisLineShow:false,showLegend:false,barGapWidthPct:60,valAxisTitle:e.unit||undefined,showValAxisTitle:!!e.unit,valAxisTitleColor:e.muted,...(e.unit&&e.unit.length<=4&&!/["\\]/.test(e.unit)?{dataLabelFormatCode:'General"'+e.unit+'"'}:{})});
      else slide.addText(e.lines.join('\n'),{...g,fontFace:e.face,fontSize:e.fontSize*.75,bold:e.bold,italic:e.italic,color:e.color,transparency:e.alpha||0,align:e.align,margin:0,paraSpaceAfter:0,lineSpacingMultiple:e.lh,valign:'top'});
    }
  }
  const raw=Buffer.from(await ppt.write({outputType:'nodebuffer'}));
  const buffer=await addMotion(raw,plan.pages.map(p=>({transition:p.transition,items:p.elements.filter(e=>e.anim).slice(0,10).map(e=>({name:e.name,effect:e.anim.effect,delay:e.anim.delay,dur:e.anim.dur}))})));
  return {buffer,plan};
}
const spawnOpen=(cmd,args,msg)=>new Promise((resolve,reject)=>{const child=spawn(cmd,args,{detached:true,stdio:'ignore'});child.once('error',()=>reject(Error(msg)));child.once('exit',code=>code?reject(Error(msg)):resolve());child.once('spawn',()=>child.unref());});
export async function openStudioInWps(filename){
  // macOS：用 WPS（com.kingsoft.wpsoffice.mac）打开；没装 WPS 时交给下面用系统默认程序打开。
  if(process.platform==='darwin')return spawnOpen('open',['-b','com.kingsoft.wpsoffice.mac',filename],'没有找到 Mac 版 WPS。');
  if(process.platform!=='win32')throw Error('当前环境无法直接启动 WPS，请下载 PPTX 后打开。');
  let output;try{output=(await execFileAsync('reg.exe',['query','HKCU\\Software\\Kingsoft\\Office\\6.0\\Common','/v','InstallRoot'],{windowsHide:true,timeout:5000})).stdout;}catch{throw Error('未找到 WPS 安装信息，请下载文件后手动打开。');}
  const match=output.match(/InstallRoot\s+REG_(?:SZ|EXPAND_SZ)\s+(.+)/i);if(!match)throw Error('无法读取 WPS 安装位置，请下载后打开。');
  const program=path.join(match[1].trim(),'office6','wpp.exe');try{await fs.access(program);}catch{throw Error('未找到 WPS 演示，请下载文件后手动打开。');}
  await new Promise((resolve,reject)=>{const child=spawn(program,[filename],{detached:true,stdio:'ignore',windowsHide:false});child.once('error',()=>reject(Error('WPS 启动失败，已生成的文件仍保留在本机。')));child.once('spawn',()=>{child.unref();resolve();});});
}

// 打开生成的 PPT：优先用 WPS 演示；没装 WPS（例如只用 Microsoft Office）时用系统默认程序打开。
export async function openStudioFile(filename){
  try{return await openStudioInWps(filename);}
  catch(e){if(process.platform==='darwin')return spawnOpen('open',[filename],'没能打开 PPT，文件已保存在本机。');if(process.platform!=='win32')throw e;
    await new Promise((resolve,reject)=>{const child=spawn('explorer.exe',[filename],{detached:true,stdio:'ignore',windowsHide:false});child.once('error',()=>reject(Error('没能打开 PPT，文件已保存在本机。')));child.once('spawn',()=>{child.unref();resolve();});});}
}
