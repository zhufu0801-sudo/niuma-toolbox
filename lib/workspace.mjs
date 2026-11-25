import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import ExcelJS from 'exceljs';
import { Document, Paragraph, Packer, Table, TableRow, TableCell } from 'docx';
import pptxgen from 'pptxgenjs';
import PDFDocument from 'pdfkit';
import fs from 'node:fs';
import path from 'node:path';
import { findCjkFont } from './fonts.mjs';
import { ensurePdfEnv } from './pdf-env.mjs';

export const formats = ['docx', 'pptx', 'xlsx', 'pdf'];
const parser = new XMLParser({ preserveOrder: true, ignoreAttributes: false, parseTagValue:false });
function descendants(nodes, tag) {
  const result = [];
  for (const node of nodes || []) for (const [key, value] of Object.entries(node)) {
    if (key === tag) result.push(value);
    if (Array.isArray(value)) result.push(...descendants(value, tag));
  }
  return result;
}
function xmlText(nodes, tag) {
  return descendants(nodes, tag).map(n => n.map(v => v['#text'] ?? '').join('')).join('');
}
function findNodes(nodes,tag){const result=[];for(const n of nodes || [])for(const [k,v] of Object.entries(n)){if(k===tag)result.push(n);if(Array.isArray(v))result.push(...findNodes(v,tag));}return result;}
function wordText(nodes){return (nodes || []).map(node=>Object.entries(node).map(([key,value])=>key==='w:t'?value.map(n=>n['#text']??'').join(''):key==='w:br'?'\n':key==='w:tab'?'\t':Array.isArray(value)?wordText(value):'').join('')).join('');}
function unwrap(nodes){return (nodes || []).flatMap(node=>node['w:sdt']?unwrap(descendants(node['w:sdt'],'w:sdtContent')[0] || []):node['w:customXml']?unwrap(node['w:customXml']):[node]);}
function textPages(text){const pages=[];let current='';for(const line of text.split('\n')){const pieces=line.match(/[^。！？.!?]+[。！？.!?]?|[。！？.!?]/g)||[''];for(let piece of pieces){while(piece.length>500){if(current){pages.push(current);current='';}const chars=Array.from(piece);pages.push(chars.slice(0,500).join(''));piece=chars.slice(500).join('');}if(current.length+piece.length>500||current.split('\n').length>=14){pages.push(current);current='';}current+=piece;}current+='\n';}if(current.trim())pages.push(current);return pages;}
export function validateBlocks(input) {
  if (!Array.isArray(input) || !input.length) throw new Error('没有可导出的内容，请先导入文件或添加内容。');
  if (input.length > 100) throw new Error('一次最多处理 100 个内容块，请拆分文件。');
  let count = 0;
  const blocks = input.map(b => {
    if (b.type === 'text') { const text = String(b.text ?? ''); count += text.length; return { type: 'text', text }; }
    if (b.type !== 'table' || !Array.isArray(b.rows) || !b.rows.length || b.rows.length > 200) throw new Error('表格需要 1～200 行。');
    const rows = b.rows.map(row => {
      if (!Array.isArray(row) || !row.length || row.length > 30) throw new Error('表格需要 1～30 列。');
      return row.map(cell => { const s = String(cell ?? ''); if (s.length > 2000) throw new Error('单元格内容超过 2000 字，请拆分。'); count += s.length; return s; });
    });
    const width = Math.max(...rows.map(r => r.length));
    return { type: 'table', rows: rows.map(r => [...r, ...Array(width-r.length).fill('')]) };
  });
  if (count > 100000) throw new Error('内容超过 10 万字，请拆分文件。');
  if (!blocks.some(b => b.type === 'text' ? b.text.trim() : b.rows.some(r => r.some(c => c.trim())))) throw new Error('内容为空，无法转换。');
  return blocks;
}
export function preflight(input, target) {
  if (!formats.includes(target)) throw new Error('请选择 PDF、Word、PPT 或 Excel。');
  const blocks = validateBlocks(input);
  if(target==='xlsx' && blocks.some(b=>b.type==='text' && b.text.split('\n').some(line=>line.length>32767))) throw new Error('Excel 单元格最多容纳 32767 个字符，请给过长文字换行或拆成多个内容块。');
  const warnings = ['这是内容重组导出：不会保留原文件的完整排版、图片、批注、公式或动画。'];
  if (target === 'pptx') warnings.push('文字按长度分页，表格按行拆页；长单元格仍可能超出版面，请打开后检查。');
  if (target === 'xlsx') warnings.push('每个内容块放入一个新工作表；数值也按文字保存，不保留原工作表名或公式。需要计算时请继续使用原始 Excel。');
  if (target === 'pdf') { if(blocks.some(b=>b.type==='table'&&b.rows[0].length>10))throw new Error('PDF 竖版最多支持 10 列表格，请拆分列或改用 Excel。'); warnings.push('表格会重建为普通网格，不保留原表格版式；导出后请检查分页。'); }
  return { blocks, warnings, target };
}
async function officeZip(buffer) {
  if (buffer[0] !== 0x50 || buffer[1] !== 0x4b) throw new Error('文件内容与扩展名不符，可能已加密或损坏。请先在 WPS 中另存为未加密文件。');
  const zip = await JSZip.loadAsync(buffer);
  const entries = Object.values(zip.files);
  if (entries.length > 5000 || entries.reduce((n,f) => n + (f._data?.uncompressedSize || 0),0) > 60 * 1024 * 1024) throw new Error('文件解压后过大，请拆分后再导入。');
  return zip;
}
export async function importFile(name, buffer) {
  const ext = String(name).split('.').pop().toLowerCase();
  if (!formats.includes(ext)) throw new Error('当前支持 .docx、.pptx、.xlsx、.pdf。旧版 doc/ppt/xls 请先在 WPS 中另存为新版格式。');
  if (!buffer.length || buffer.length > 15*1024*1024) throw new Error('文件不能为空，且不能超过 15 MB。');
  let blocks = [], warnings = ['仅提取文字和静态表格，不读取图片、批注、公式逻辑或动画；原文件保持不变。'];
  if (ext === 'pdf') {
    if (buffer.subarray(0,5).toString() !== '%PDF-') throw new Error('这不是有效的 PDF 文件。');
    await ensurePdfEnv();
    const { PDFParse } = await import('pdf-parse');
    const pdf = new PDFParse({data:new Uint8Array(buffer)});
    try { const data = await pdf.getText(); blocks = data.pages.map(p => ({type:'text', text:p.text})).filter(b => b.text.trim()); }
    finally { await pdf.destroy(); }
    warnings.push('PDF 按页提取文字，阅读顺序和表格结构可能变化；扫描件暂不支持 OCR。');
    if (!blocks.length) throw new Error('PDF 中未找到可提取文字，可能是扫描件。请先做 OCR 识别。');
  } else {
    const zip = await officeZip(buffer);
    if (ext === 'docx') {
      const entry = zip.file('word/document.xml');
      if (!entry) throw new Error('文件不是有效的 Word 文档。');
      const body = descendants(parser.parse(await entry.async('string')), 'w:body')[0] || [];
      for (const node of unwrap(body)) {
        if (node['w:p']) { const text=wordText(node['w:p']); if(text.trim()){const previous=blocks.at(-1);if(previous?.type==='text')previous.text+='\n'+text;else blocks.push({type:'text',text});} }
        if (node['w:tbl']) {
          const rows=unwrap(node['w:tbl']).filter(n=>n['w:tr']).map(node=>{
            const row=node['w:tr'];const output=[];
            const before=findNodes(row,'w:gridBefore')[0]?.[':@']?.['@_w:val'];for(let i=0;i<Math.min(Number(before)||0,30);i++)output.push('');
            for(const cellNode of unwrap(row).filter(n=>n['w:tc'])){const cell=cellNode['w:tc'];const span=Number(findNodes(cell,'w:gridSpan')[0]?.[':@']?.['@_w:val'] || 1);if(!Number.isInteger(span)||span<1||span>30)throw new Error('Word 表格合并范围不支持，请拆分合并单元格后重试。');output.push(descendants(cell,'w:p').map(p=>wordText(p)).join('\n'));for(let i=1;i<span;i++)output.push('');}
            return output;
          });
          if(rows.length) blocks.push({type:'table',rows});
        }
      }
      warnings.push('Word 横向合并单元格会补空列占位；纵向合并关系不保留。页眉、页脚和文本框可能未包含。');
    } else if (ext === 'pptx') {
      const slides=Object.keys(zip.files).filter(n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a,b)=>Number(a.match(/slide(\d+)/)[1])-Number(b.match(/slide(\d+)/)[1]));
      if(!slides.length) throw new Error('文件中没有可读取的幻灯片。');
      for(const slide of slides) { const xml=parser.parse(await zip.file(slide).async('string')); const text=descendants(xml,'a:p').map(p=>xmlText(p,'a:t')).filter(Boolean).join('\n'); if(text.trim()) blocks.push({type:'text',text}); }
      warnings.push('PPT 按内部页序提取文字，表格暂按文字提取；不保留母版与布局。');
    } else {
      if(!zip.file('xl/workbook.xml')) throw new Error('文件不是有效的 Excel 工作簿。');
      const workbook=new ExcelJS.Workbook(); await workbook.xlsx.load(buffer);
      for (const sheet of workbook.worksheets) {
        if (sheet.rowCount > 200 || sheet.columnCount > 30) throw new Error(`工作表“${sheet.name}”超过 200 行或 30 列，请缩小数据范围。`);
        const rows=[];
        sheet.eachRow({includeEmpty:true}, row => { const cells=[]; for(let i=1;i<=sheet.columnCount;i++) { const cell=row.getCell(i); cells.push(cell.text || ''); } if(cells.some(Boolean)) rows.push(cells); });
        if(rows.length) blocks.push({type:'table',rows});
      }
      warnings.push('Excel 读取单元格显示值；公式无缓存结果时可能无法提取，不保留合并关系。');
    }
  }
  return { blocks:validateBlocks(blocks), warnings, source:ext };
}
export async function exportFile(input, target) {
  const {blocks}=preflight(input,target);
  if(target==='docx') return Packer.toBuffer(new Document({sections:[{children:blocks.flatMap(b => b.type==='text' ? b.text.split('\n').map(text=>new Paragraph(text)) : [new Table({rows:b.rows.map(row=>new TableRow({children:row.map(text=>new TableCell({children:[new Paragraph(text)]}))}))})])}]}));
  if(target==='xlsx') {
    const wb=new ExcelJS.Workbook();
    blocks.forEach((b,i)=>{const s=wb.addWorksheet(`内容${i+1}`); s.addRows(b.type==='table'?b.rows:b.text.split('\n').map(t=>[t])); s.columns.forEach(c=>{c.width=30; c.alignment={wrapText:true,vertical:'top'};}); });
    return Buffer.from(await wb.xlsx.writeBuffer());
  }
  if(target==='pptx') {
    const ppt=new pptxgen(); ppt.layout='LAYOUT_WIDE'; ppt.author='WPS AI 内容工作台';
    for(const b of blocks) {
      if(b.type==='text') for(const page of textPages(b.text)) {const slide=ppt.addSlide();slide.addText(page,{x:.6,y:.5,w:12.1,h:6.4,fontFace:'Microsoft YaHei',fontSize:20,fit:'shrink',valign:'top'});}
      else {const header=b.rows[0],data=b.rows.slice(1);for(let i=0;i<Math.max(data.length,1);i+=11) {const slide=ppt.addSlide(),rows=[header,...data.slice(i,i+11)];slide.addTable(rows,{x:.4,y:.4,w:12.5,h:Math.min(6.5,rows.length*.5),fontFace:'Microsoft YaHei',fontSize:14,border:{pt:1,color:'CAD8D1'},color:'243B34',margin:5,autoPage:false});}}
    }
    return Buffer.from(await ppt.write({outputType:'nodebuffer'}));
  }
  const found=findCjkFont();
  if(!found) throw new Error('缺少中文 PDF 字体，请安装中文字体或设置 COACH_PDF_FONT 后重试。');
  return new Promise((resolve,reject)=>{
    const pdf=new PDFDocument({margin:45,size:'A4'}), chunks=[];
    pdf.on('data',c=>chunks.push(c)); pdf.on('end',()=>resolve(Buffer.concat(chunks)));pdf.on('error',reject);
    try {pdf.font(found.file, found.family).fontSize(11);for(const b of blocks){if(b.type==='text')pdf.text(b.text,{lineGap:4});else{const width=(pdf.page.width-90)/b.rows[0].length;for(const row of b.rows)for(const cell of row)if(pdf.heightOfString(cell,{width:width-12})>600)throw new Error('PDF 表格单元格过长，无法完整放入一页，请拆分内容或改用 Word。');pdf.table({columnStyles:b.rows[0].map(()=>width),data:b.rows,defaultStyle:{border:0.5,padding:5}});}pdf.moveDown();}pdf.end();}catch(e){pdf.destroy();reject(e);}
  });
}
