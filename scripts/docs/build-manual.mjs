// Run from the repository: node scripts/docs/build-manual.mjs [--out release/docs]
// Uses existing docx/pdfkit dependencies; never downloads fonts or packages.
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
import PDFDocument from 'pdfkit';
import {Document,Packer,Paragraph,TextRun,HeadingLevel,Table,TableRow,TableCell,WidthType,BorderStyle,ShadingType,TableOfContents,Header,Footer,PageNumber,Bookmark,InternalHyperlink,AlignmentType} from 'docx';
const require=createRequire(import.meta.url),fontkit=require('fontkit');
const ROOT=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'../..');
const ai=process.argv.indexOf('--out'),OUT=path.resolve(ai>=0?process.argv[ai+1]:path.join(ROOT,'release/docs'));
fs.mkdirSync(OUT,{recursive:true});
const FONTDIR=process.env.WINDIR?path.join(process.env.WINDIR,'Fonts'):'';
import {fontSpec,MISSING_FONT_EXIT} from './manual-fonts.mjs';
import checks from './manual-checks.cjs';
const title={zh:'办公工具箱详细使用说明书',ja:'オフィスツールボックス 詳細操作ガイド',en:'Office Toolbox Detailed User Manual'};
const tocTitle={zh:'自动目录',ja:'自動目次',en:'Automatic contents'};
const strip=s=>s.replace(/<a id="[^"]+"><\/a>/g,'').replace(/\[([^\]]+)\]\([^)]+\)/g,'$1').replace(/\*\*([^*]+)\*\*/g,'$1').replace(/`([^`]+)`/g,'$1').replace(/<br\s*\/?\s*>/g,'\n');
function parse(md){const lines=md.split(/\r?\n/),blocks=[];let anchor=null;
 for(let i=0;i<lines.length;){let s=lines[i].trim();if(!s){i++;continue;}const a=s.match(/^<a id="([^"]+)"><\/a>$/);if(a){anchor=a[1];i++;continue;}
  const h=s.match(/^(#{1,4}) (.+)/);if(h){blocks.push({type:'heading',level:h[1].length,text:strip(h[2]),anchor});anchor=null;i++;continue;}
  if(s.startsWith('|')&&lines[i+1]?.match(/^\|[- :|]+\|$/)){const rows=[];while(i<lines.length&&lines[i].trim().startsWith('|')){const row=lines[i++].trim();if(/^\|[- :|]+\|$/.test(row))continue;rows.push(row.slice(1,-1).split('|').map(v=>({text:strip(v.trim()),anchor:v.match(/<a id="([^"]+)/)?.[1]})));}blocks.push({type:'table',rows});continue;}
  if(/^\d+\. /.test(s)||/^[-*] /.test(s)){blocks.push({type:'list',text:strip(s),target:s.match(/\]\(#([^)]*)\)/)?.[1]});i++;continue;}
  if(/^\*\*[^*]+\*\*$/.test(s)){blocks.push({type:'label',text:strip(s)});i++;continue;}
  let p=[s];i++;while(i<lines.length&&lines[i].trim()&&!/^(?:#|<a |\||\*\*|\d+\. |[-*] )/.test(lines[i].trim()))p.push(lines[i++].trim());blocks.push({type:'para',text:strip(p.join(' '))});
 }return blocks;}
function symbolSpec(fallback){for(const file of [path.join(FONTDIR,'seguisym.ttf'),'/System/Library/Fonts/Apple Symbols.ttf']){if(fs.existsSync(file)){const font=fontkit.openSync(file);return {file,face:undefined,name:font.familyName,font};}}return fallback;}
function fontRuns(text,fonts,bold=false){const a=bold?fonts.bold:fonts.body,b=bold?fonts.fallbackBold:fonts.fallback;let chunks=[],cur='',last=null;for(const c of text){const cp=c.codePointAt(0),f=cp<32?a:[a,b,fonts.symbols].find(f=>f.font.hasGlyphForCodePoint(cp));if(!f)throw Error('Missing glyph '+c+' U+'+cp.toString(16));if(last&&f!==last){chunks.push({text:cur,spec:last});cur='';}cur+=c;last=f;}if(cur)chunks.push({text:cur,spec:last});return chunks;}
function wordRuns(text,fonts,bold=false){return fontRuns(text,fonts,bold).map(r=>new TextRun({text:r.text,font:{name:r.spec.name,eastAsia:r.spec.name},bold,color:'000000'}));}
function wordPara(block,fonts){let children=wordRuns(block.text,fonts,block.type==='label');if(block.anchor)children=[new Bookmark({id:block.anchor,children})];if(block.target)children=[new InternalHyperlink({anchor:block.target,children:wordRuns(block.text,fonts)})];return new Paragraph({children,spacing:{after:block.type==='label'?50:35,line:250},keepNext:block.type==='label',widowControl:true});}
const border={style:BorderStyle.SINGLE,size:4,color:'D9D9D9'};
async function makeWord(blocks,lang,fonts){const children=[];let tocDone=false;
 for(const b of blocks){if(b.type==='heading'){if(b.level===1){children.push(new Paragraph({text:b.text,heading:HeadingLevel.TITLE}));continue;}
  if(!b.anchor&&['目录','目次','Contents'].includes(b.text)){children.push(new TableOfContents(tocTitle[lang],{hyperlink:true,headingStyleRange:'1-1'}));tocDone=true;continue;}
  const h=b.level===2?HeadingLevel.HEADING_1:HeadingLevel.HEADING_2;const r=wordRuns(b.text,fonts,true);children.push(new Paragraph({heading:h,children:b.anchor?[new Bookmark({id:b.anchor,children:r})]:r,pageBreakBefore:b.level===2,keepNext:true}));continue;}
  if(b.type==='list'&&b.target&&tocDone&&!children.some(x=>x.__body)){
   // Keep linked chapter navigation as a usable fallback before fields are refreshed.
   continue;
  }
  if(b.type==='table'){const cols=b.rows[0].length,widths=cols===2?[3300,6338]:cols===3?[2400,3500,3738]:Array(cols).fill(Math.floor(9638/cols));children.push(new Table({width:{size:9638,type:WidthType.DXA},columnWidths:widths,rows:b.rows.map((row,i)=>new TableRow({tableHeader:i===0,cantSplit:false,children:row.map((cell,j)=>new TableCell({width:{size:widths[j],type:WidthType.DXA},margins:{top:60,bottom:60,left:100,right:100},borders:{top:border,bottom:border,left:border,right:border},shading:{type:ShadingType.CLEAR,fill:i===0?'E8EEF4':i%2?'FFFFFF':'F7F8FA'},children:[new Paragraph({children:cell.anchor?[new Bookmark({id:cell.anchor,children:wordRuns(cell.text,fonts,i===0)})]:wordRuns(cell.text,fonts,i===0),spacing:{after:20,line:250},widowControl:true})]}))}))}));children.push(new Paragraph({text:'',spacing:{after:80}}));continue;}
  children.push(wordPara(b,fonts));
 }
 const font={name:fonts.body.name,eastAsia:fonts.body.name};
 const doc=new Document({
  creator:'Office Toolbox',title:title[lang],description:'Detailed manual, source 2026-11-25',features:{updateFields:true},
  styles:{
   default:{document:{run:{font,size:22,color:'000000'},paragraph:{spacing:{after:35,line:250}}}},
   paragraphStyles:[
    {id:'Title',name:'Title',basedOn:'Normal',next:'Normal',run:{font,size:42,bold:true,color:'000000'},paragraph:{spacing:{before:240,after:320}}},
    {id:'Heading1',name:'Heading 1',basedOn:'Normal',next:'Normal',quickFormat:true,run:{font,size:32,bold:true,color:'000000'},paragraph:{spacing:{before:200,after:180},keepNext:true,outlineLevel:0}},
    {id:'Heading2',name:'Heading 2',basedOn:'Normal',next:'Normal',quickFormat:true,run:{font,size:26,bold:true,color:'000000'},paragraph:{spacing:{before:140,after:90},keepNext:true,outlineLevel:1}}
   ]
  },
  sections:[{
   properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1134,right:1134}}},
   headers:{default:new Header({children:[new Paragraph({children:[new TextRun({text:title[lang],font,size:16,color:'444444'})],spacing:{after:40}})]})},
   footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({children:[PageNumber.CURRENT],font,size:18})]})]})},children
  }]
 });
 await fs.promises.writeFile(path.join(OUT,'manual-'+lang+'.docx'),await Packer.toBuffer(doc));
}
async function makePdf(blocks,lang,fonts){const file=path.join(OUT,'manual-'+lang+'.pdf');const pdf=new PDFDocument({size:'A4',margins:{top:56.7,bottom:56.7,left:56.7,right:56.7},bufferPages:true,info:{Title:title[lang],Author:'Office Toolbox',Subject:'Detailed manual 2026-11-25'}});const stream=fs.createWriteStream(file);pdf.pipe(stream);let p=0;for(const [key,spec]of Object.entries(fonts))pdf.registerFont(key,spec.file,spec.face);
 const W=pdf.page.width-113.4,top=56.7,maxY=pdf.page.height-56.7,rowsMeta=[],outlineStack=[];
 function runs(text,bold=false){return fontRuns(text,fonts,bold).map(r=>({text:r.text,font:Object.entries(fonts).find(([,s])=>s===r.spec)?.[0]||'body'}));}
 function measure(text,width,size=11,bold=false){pdf.font(bold?'bold':'body').fontSize(size);return pdf.heightOfString(text,{width,lineGap:1.2});}
 function next(){pdf.addPage();p++;}
 function ensure(h){if(pdf.y+h>maxY)next();}
 function draw(text,x,y,width,size=11,bold=false,options={}){const rr=runs(text,bold);pdf.fontSize(size);rr.forEach((r,i)=>{pdf.font(r.font);pdf.text(r.text,i===0?x:undefined,i===0?y:undefined,{width,lineGap:1.2,continued:i<rr.length-1,...options});});return pdf.y;}
 for(const b of blocks){
  if(b.type==='heading'){const size=b.level===1?23:b.level===2?17:13;
   if(b.level===2&&b.anchor&&pdf.y>top+20)next();else ensure(measure(b.text,W,size,true)+75);
   if(b.anchor)pdf.addNamedDestination(b.anchor);const oy=pdf.y;draw(b.text,56.7,oy,W,size,true);pdf.y+=b.level===1?20:12;
   if(b.anchor){if(b.level===2){const o=pdf.outline.addItem(b.text);outlineStack.length=0;outlineStack.push(o);}else (outlineStack[0]||pdf.outline).addItem(b.text);}
   continue;
  }
  if(b.type==='table'){const n=b.rows[0].length,ws=n===2?[W*.34,W*.66]:n===3?[W*.25,W*.36,W*.39]:Array(n).fill(W/n);let first=true;
   function row(cells,isHead=false){const heights=cells.map((c,i)=>measure(c.text,ws[i]-12,11,isHead)+16),h=Math.max(...heights);if(h>maxY-top)throw Error('Table row exceeds page: '+cells[0].text.slice(0,80));if(pdf.y+h>maxY){next();if(!isHead)row(b.rows[0],true);}let x=56.7,y=pdf.y;cells.forEach((c,i)=>{pdf.save().rect(x,y,ws[i],h).fillAndStroke(isHead?'#E8EEF4':'#FFFFFF','#D9D9D9').restore();if(c.anchor)pdf.addNamedDestination(c.anchor);draw(c.text,x+6,y+7,ws[i]-12,11,isHead);x+=ws[i];});rowsMeta.push({page:p+1,y,height:h});pdf.y=y+h;}
   for(const cells of b.rows)row(cells,first?(first=false,true):false);pdf.y+=12;continue;
  }
  const bold=b.type==='label',h=measure(b.text,W,11,bold);ensure(h+(bold?36:0)+7);draw(b.text,56.7,pdf.y,W,11,bold,b.target?{goTo:b.target}:{});pdf.y+=bold?3:5;
 }
 const range=pdf.bufferedPageRange();for(let i=range.start;i<range.start+range.count;i++){pdf.switchToPage(i);const bottom=pdf.page.margins.bottom;pdf.page.margins.bottom=0;pdf.font('body').fontSize(8).fillColor('#555555').text(title[lang],56.7,29,{width:W,lineBreak:false});pdf.text(String(i+1),56.7,pdf.page.height-35,{width:W,align:'center',lineBreak:false});pdf.page.margins.bottom=bottom;}
 if(pdf.bufferedPageRange().count!==range.count)throw Error('Header/footer unexpectedly added pages');
 pdf.end();await new Promise((resolve,reject)=>{stream.on('finish',resolve);stream.on('error',reject);});return {pages:range.count,tables:rowsMeta.length,fonts:Object.fromEntries(Object.entries(fonts).map(([k,v])=>[k,{file:v.file,face:v.face}]))};
}
async function main(){
const result={builtAt:new Date().toISOString(),languages:{}};
for(const lang of ['zh','ja','en']){const md=fs.readFileSync(path.join(ROOT,'docs/manual',lang+'.md'),'utf8'),blocks=parse(md),body=fontSpec(lang),bold=fontSpec(lang,true),fonts={body,bold,fallback:fontSpec('zh'),fallbackBold:fontSpec('zh',true),symbols:symbolSpec(body)};
 // Check source before producing either format.
 if(lang!=='en')checks.validateManual(md,lang);
 // Verify glyph coverage before producing either format.
 for(const b of blocks){if(b.text)fontRuns(b.text,fonts);for(const r of b.rows||[])for(const c of r)fontRuns(c.text,fonts);}
 await makeWord(blocks,lang,fonts);result.languages[lang]=await makePdf(blocks,lang,fonts);console.log(lang,result.languages[lang].pages+' PDF pages');}
fs.writeFileSync(path.join(OUT,'build-report.json'),JSON.stringify(result,null,2));

}
main().catch(error=>{console.error(error.message);process.exitCode=error.code==='MANUAL_FONT_MISSING'?MISSING_FONT_EXIT:1;});
