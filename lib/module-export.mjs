import fs from 'node:fs/promises';import path from 'node:path';import {randomBytes} from 'node:crypto';
import {Document,Paragraph,TextRun,Packer,Table,TableRow,TableCell,HeadingLevel,WidthType} from 'docx';
import ExcelJS from 'exceljs';import pptxgen from 'pptxgenjs';
import {checkCardBlock} from './cards.mjs';import {checkModuleFormat} from './module-format.mjs';import {exportFile} from './workspace.mjs';
const align=['left','center','right','justify'];const color=value=>value?.replace('#','');
function textStyle(s={}){return {font:s.font||'宋体',size:(s.size||12)*2,bold:s.bold,italics:s.italic,color:color(s.color)};}
function paragraph(text,s={}){return new Paragraph({heading:s.level?HeadingLevel['HEADING_'+Math.min(6,s.level)]:undefined,alignment:align[s.alignment],children:[new TextRun({text,...textStyle(s)})]});}
function pagesOf(block,format){const lines=block.text.split('\n'),paragraphs=format?.paragraphs;if(paragraphs?.some(p=>p.level)){const pages=[];lines.forEach((line,i)=>{if(paragraphs[i]?.level||!pages.length)pages.push({title:line,body:[]});else pages.at(-1).body.push(line);});return pages;}return block.text.split(/\n\s*\n/).filter(p=>p.trim()).map(page=>{const [title,...body]=page.split('\n');return {title,body};});}
export async function exportModule(card,target){const block=checkCardBlock(card.block),format=checkModuleFormat(card.format,block),base=format?.base||{};
 if(!['docx','xlsx','pptx','pdf'].includes(target))throw Error('请选择 Word、Excel、PPT 或 PDF 格式。');
 if(target==='pdf')return exportFile([block],target);
 if(target==='docx'){
  let children;
  if(block.type==='text')children=block.text.split('\n').map((text,i)=>paragraph(text,{...base,...format?.paragraphs?.[i]}));
  else{
   const widths=format?.columnWidths,total=widths?.reduce((a,b)=>a+b,0),columnWidths=widths?.map(w=>Math.round(w/total*9000));
   const rows=block.rows.map((row,r)=>new TableRow({tableHeader:r===0,children:row.map((text,c)=>{
    const s={...base,bold:r===0,...format?.cells?.[r]?.[c]};
    return new TableCell({shading:s.background?{fill:color(s.background)}:undefined,children:[paragraph(text,s)]});
   })}));
   children=[new Table({width:{size:100,type:WidthType.PERCENTAGE},columnWidths,rows})];
  }
  return Packer.toBuffer(new Document({sections:[{children}]}));
 }
 if(target==='xlsx'){const wb=new ExcelJS.Workbook(),sheet=wb.addWorksheet('内容模块'),rows=block.type==='table'?block.rows:block.text.split('\n').map(text=>[text]);sheet.addRows(rows);rows.forEach((row,r)=>row.forEach((text,c)=>{const s={...base,...(block.type==='table'?format?.cells?.[r]?.[c]:format?.paragraphs?.[r])},cell=sheet.getCell(r+1,c+1);cell.font={name:s.font||'宋体',size:s.size||12,bold:s.bold??(block.type==='table'&&r===0||!!s.level),italic:s.italic,color:s.color?{argb:'FF'+color(s.color)}:undefined};cell.alignment={wrapText:true,vertical:'top',horizontal:align[s.alignment]};if(s.background)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF'+color(s.background)}};}));sheet.columns.forEach((column,c)=>{column.width=Math.max(8,Math.min(60,format?.columnWidths?.[c]?format.columnWidths[c]/6:Math.max(...rows.map(r=>String(r[c]||'').length))*1.6+3));});return Buffer.from(await wb.xlsx.writeBuffer());}
 const ppt=new pptxgen();ppt.layout='LAYOUT_WIDE';ppt.author='办公工具箱';
 if(block.type==='table'){
  for(let i=1;i<Math.max(block.rows.length,2);i+=9){
   const indexes=[0,...Array.from({length:Math.min(9,block.rows.length-i)},(_,k)=>i+k)];
   const charsPerLine=Math.max(3,Math.floor((900/block.rows[0].length-10)/14));
   const requiredHeight=indexes.reduce((height,r)=>height+Math.max(...block.rows[r].map(text=>text.split('\n').reduce((n,line)=>n+Math.max(1,Math.ceil(Array.from(line).length/charsPerLine)),0)))*.24+.15,0);
   if(requiredHeight>6.2)throw Error('表格内容过长，无法完整放入 PPT 页面。请缩短单元格文字、拆分模块，或在设置中改为保存 Excel。');
   const rows=indexes.map(r=>block.rows[r].map((text,c)=>{const s={...base,...format?.cells?.[r]?.[c]};return {text,options:{fontFace:s.font||'Microsoft YaHei',bold:s.bold??r===0,italic:s.italic,color:color(s.color)||'203B40',fill:color(s.background)||undefined,align:align[s.alignment]}};}));
   ppt.addSlide().addTable(rows,{x:.4,y:.5,w:12.5,fontSize:14,border:{pt:1,color:'C6D9DC'},margin:5,autoPage:false});
  }
 }else{
  const pages=pagesOf(block,format);if(pages.length>30)throw Error('模块超过 30 页，请拆分后保存。');
  for(const page of pages){if(page.title.length>150||page.body.join('\n').length>1200)throw Error('单页内容过长，请拆分模块后保存 PPT，避免文字超出页面。');const slide=ppt.addSlide();slide.addText(page.title,{x:.5,y:.35,w:12.3,h:.8,fontFace:base.font||'Microsoft YaHei',fontSize:28,bold:true,breakLine:false,fit:'shrink'});if(page.body.length)slide.addText(page.body.join('\n'),{x:.5,y:1.4,w:12.3,h:5.5,fontFace:base.font||'Microsoft YaHei',fontSize:20,fit:'shrink',valign:'top'});}
 }
 return Buffer.from(await ppt.write({outputType:'nodebuffer'}));
}
export async function saveModuleToDesktop(card,target,desktop){if(!desktop||!path.isAbsolute(desktop))throw Error('无法确认桌面位置，请检查 Windows 桌面文件夹设置。');const buffer=await exportModule(card,target);await fs.mkdir(desktop,{recursive:true});let name=String(card.title||'内容模块').replace(/[<>:"/\\|?*\x00-\x1f]/g,'_').replace(/[. ]+$/,'').slice(0,55)||'内容模块';if(/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name))name='模块-'+name;const stamp=new Date().toISOString().replace(/[-:TZ.]/g,'');for(let n=0;n<4;n++){const targetPath=path.join(desktop,name+'-'+stamp+'-'+randomBytes(3).toString('hex')+'.'+target);try{await fs.writeFile(targetPath,buffer,{flag:'wx'});return {filePath:targetPath,format:target,warnings:target==='pdf'?['PDF 会重建内容，请检查版式。']:['保留已读取的文字和表格样式；图片、公式、合并关系及完整原排版不包含。']};}catch(e){if(e.code!=='EEXIST')throw Error('无法保存到桌面：'+e.message);}}throw Error('无法生成唯一文件名，请重试。');}
