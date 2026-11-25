// 按 Excel 名单批量生成文档（2026-11-21）：Word 模板里写 {{姓名}}、{姓名}、《姓名》 或 «姓名»（括号里是名单的列名），
// 每一行名单生成一份。可以合成一个文档（每份之间分页，方便一次打印），也可以每行存一个文件（文件名用指定的列）。
// 只替换名单里真有的列名，正文里其他的《书名》、{…} 不会动。WPS 把一个占位符拆成几段文字时也能认出来。
import fs from 'node:fs/promises';import path from 'node:path';
import JSZip from 'jszip';import ExcelJS from 'exceljs';

const ESC={'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'};
const esc=s=>String(s).replace(/[&<>"]/g,c=>ESC[c]);
const unesc=s=>String(s).replace(/&(lt|gt|quot|apos|amp|#\d+|#x[0-9a-f]+);/gi,(m,e)=>e==='lt'?'<':e==='gt'?'>':e==='quot'?'"':e==='apos'?"'":e==='amp'?'&':e[1]==='x'||e[1]==='X'?String.fromCodePoint(parseInt(e.slice(2),16)):String.fromCodePoint(Number(e.slice(1))));
export const MAX_ROWS=500;

// ExcelJS 的日期格子 .text 是英文长串，这里按格子的数字格式自己写成 Excel 里看到的样子（ExcelJS 的日期按 UTC 存）。
function dateText(d,fmt){if(isNaN(d))return '';const f=String(fmt||'').split(';')[0].replace(/\[[^\]]*\]/g,'');const Y=d.getUTCFullYear(),M=d.getUTCMonth()+1,D=d.getUTCDate(),h=d.getUTCHours(),mi=d.getUTCMinutes(),p2=n=>String(n).padStart(2,'0');
  if(!/[yd]/i.test(f)&&!/m/i.test(f))return Y+'/'+M+'/'+D;let out='',i=0,lastH=false;
  while(i<f.length){const c=f[i];if(c==='"'){const j=f.indexOf('"',i+1);out+=f.slice(i+1,j<0?f.length:j);i=j<0?f.length:j+1;continue;}if(c==='\\'){out+=f[i+1]||'';i+=2;continue;}
    const run=(f.slice(i).match(/^(y+|m+|d+|h+|s+)/i)||[''])[0];if(!run){out+=c;i++;continue;}const k=run[0].toLowerCase(),n=run.length;i+=n;
    if(k==='y')out+=n<=2?p2(Y%100):Y;else if(k==='d')out+=n>=4?['日','月','火','水','木','金','土'][d.getUTCDay()]:n===3?['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][d.getUTCDay()]:n===2?p2(D):D;
    else if(k==='h'){out+=n>=2?p2(h):h;lastH=true;continue;}else if(k==='s')out+=p2(d.getUTCSeconds());else out+=lastH?p2(mi):n>=2?p2(M):M;lastH=false;}
  return out.trim()||Y+'/'+M+'/'+D;}
function cellText(cell){const v=cell.value;if(v instanceof Date)return dateText(v,cell.numFmt);if(v&&typeof v==='object'&&v.result instanceof Date)return dateText(v.result,cell.numFmt);return String(cell.text??'');}

// 名单：第一个工作表，第一行有内容的行是表头；按 Excel 里显示的样子取值（日期、金额保持原来的格式）。
export async function readList(file){const wb=new ExcelJS.Workbook();try{await wb.xlsx.readFile(file);}catch{throw Error('名单打不开。请用 .xlsx 格式（WPS 表格里“另存为 Excel 工作簿”）。');}
  const ws=wb.worksheets[0];if(!ws)throw Error('名单里没有工作表。');let headerRow=0;for(let r=1;r<=Math.min(ws.rowCount,20);r++){if(ws.getRow(r).values.some(v=>v!=null&&String(v).trim()!=='')){headerRow=r;break;}}
  if(!headerRow)throw Error('名单是空的。');const head=ws.getRow(headerRow),columns=[];head.eachCell({includeEmpty:false},(c,i)=>{const name=String(c.text||'').trim();if(name)columns.push({name,col:i});});
  if(!columns.length)throw Error('名单第一行要写列名，例如：姓名、部门。');const rows=[];
  for(let r=headerRow+1;r<=ws.rowCount;r++){const row=ws.getRow(r),rec={};let any=false;for(const c of columns){const t=cellText(row.getCell(c.col)).trim();rec[c.name]=t;if(t)any=true;}if(any)rows.push(rec);if(rows.length>MAX_ROWS)throw Error('名单超过 '+MAX_ROWS+' 行，请分成几份。');}
  if(!rows.length)throw Error('名单只有表头，没有数据行。');return {columns:columns.map(c=>c.name),rows};}

const OPEN=['{{','{','《','«'],CLOSE={'{{':'}}','{':'}','《':'》','«':'»'};
function placeholderRe(columns){const names=columns.slice().sort((a,b)=>b.length-a.length).map(n=>n.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')).join('|');return new RegExp('\\{\\{\\s*('+names+')\\s*\\}\\}|\\{\\s*('+names+')\\s*\\}|《\\s*('+names+')\\s*》|«\\s*('+names+')\\s*»','g');}
const ANY_PLACEHOLDER=/\{\{\s*([^{}\s][^{}]{0,30}?)\s*\}\}/g;

// 在一个段落里替换：先把各段 <w:t> 的文字连起来找，再把结果写回第一段，后面被占用的部分删掉。
function fillParagraph(p,re,record,found){const parts=[];const tre=/(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)|<w:t\s*\/>/g;let m;while((m=tre.exec(p))){if(!m[1])continue;parts.push({start:m.index,len:m[0].length,open:m[1],text:unesc(m[2])});}
  if(!parts.length)return p;const all=parts.map(x=>x.text).join('');re.lastIndex=0;const hits=[];while((m=re.exec(all))){const name=m[1]||m[2]||m[3]||m[4];hits.push({at:m.index,len:m[0].length,value:record?String(record[name]??''):m[0],name});found&&found.add(name);}
  if(!hits.length||!record)return p;
  const offs=[];let acc=0;for(const x of parts){offs.push(acc);acc+=x.text.length;}
  const locate=pos=>{for(let i=parts.length-1;i>=0;i--)if(pos>=offs[i])return i;return 0;};
  for(let h=hits.length-1;h>=0;h--){const {at,len,value}=hits[h],s=locate(at),e=locate(at+len-1),so=at-offs[s],eo=at+len-offs[e];
    if(s===e){parts[s].text=parts[s].text.slice(0,so)+value+parts[s].text.slice(eo);}
    else{parts[e].text=parts[e].text.slice(eo);for(let i=s+1;i<e;i++)parts[i].text='';parts[s].text=parts[s].text.slice(0,so)+value;}}
  let out='',last=0;for(const x of parts){out+=p.slice(last,x.start)+'<w:t xml:space="preserve">'+esc(x.text).replace(/\r?\n/g,'</w:t><w:br/><w:t xml:space="preserve">')+'</w:t>';last=x.start+x.len;}return out+p.slice(last);}
function fillXml(xml,re,record,found){return xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g,p=>fillParagraph(p,re,record,found));}
const PARTS=/^word\/(document|header\d*|footer\d*)\.xml$/;

// 预览：模板里用了哪些列、名单里没有的 {{…}}、名单里有但模板没用的列。
export async function inspectMerge({template,list}){const {columns,rows}=await readList(list),zip=await loadTemplate(template),re=placeholderRe(columns),found=new Set(),unknown=new Set();
  for(const name of Object.keys(zip.files).filter(n=>PARTS.test(n))){const xml=await zip.file(name).async('string');fillXml(xml,re,null,found);const plain=xml.replace(/<[^>]+>/g,'');let m;ANY_PLACEHOLDER.lastIndex=0;while((m=ANY_PLACEHOLDER.exec(plain)))if(!columns.includes(m[1].trim()))unknown.add(m[1].trim());}
  return {columns,count:rows.length,used:columns.filter(c=>found.has(c)),unused:columns.filter(c=>!found.has(c)),unknown:[...unknown],sample:rows[0]};}
async function loadTemplate(file){if(!/\.docx$/i.test(String(file)))throw Error('模板要先保存成 .docx 格式（WPS 文字里“另存为 Word 文档”）。');try{return await JSZip.loadAsync(await fs.readFile(file));}catch{throw Error('模板文件打不开，请先保存后再试。');}}

const safe=s=>String(s||'').replace(/[\\/:*?"<>|\r\n\t]+/g,' ').trim().slice(0,60);
async function freeName(dir,base,ext){let name=base+ext,i=2;while(await fs.stat(path.join(dir,name)).then(()=>true,()=>false))name=base+'-'+(i++)+ext;return path.join(dir,name);}

// mode：combined（合成一个文档，每份之间分页）| separate（每行一个文件，存进新文件夹）。
export async function runMerge({template,list,mode='combined',nameColumn='',outputDir}){const {columns,rows}=await readList(list),zip=await loadTemplate(template),re=placeholderRe(columns);
  const docXml=await zip.file('word/document.xml').async('string'),found=new Set();fillXml(docXml,re,null,found);
  if(!found.size)throw Error('模板里没有找到名单的列名。请在文档里写上 {{'+columns[0]+'}} 这样的占位符（列名要和名单第一行一致），保存后再试。');
  const base=safe(path.basename(template,path.extname(template)));await fs.mkdir(outputDir,{recursive:true});
  const otherParts=Object.keys(zip.files).filter(n=>PARTS.test(n)&&n!=='word/document.xml');
  if(mode==='separate'){if(nameColumn&&!columns.includes(nameColumn))throw Error('文件名用的列“'+nameColumn+'”不在名单里。');
    const folder=await freeName(outputDir,base+'-批量生成','');await fs.mkdir(folder);const used=new Set();
    for(let i=0;i<rows.length;i++){const z=await JSZip.loadAsync(await zip.generateAsync({type:'nodebuffer'}));z.file('word/document.xml',fillXml(docXml,re,rows[i]));for(const n of otherParts)z.file(n,fillXml(await zip.file(n).async('string'),re,rows[i]));
      let name=safe(nameColumn?rows[i][nameColumn]:'')||String(i+1).padStart(3,'0');if(used.has(name))name+='-'+(i+1);used.add(name);
      await fs.writeFile(path.join(folder,name+'.docx'),await z.generateAsync({type:'nodebuffer',compression:'DEFLATE'}));}
    return {folder,count:rows.length,used:[...found]};}
  // 合成：正文部分按行重复，中间插分页；图片等的编号按份错开，避免 Word 提示“内容有问题”。
  const body=docXml.match(/<w:body>([\s\S]*)<\/w:body>/);if(!body)throw Error('模板格式无法识别。');
  const sect=body[1].match(/<w:sectPr[\s\S]*?<\/w:sectPr>\s*$/),content=sect?body[1].slice(0,body[1].length-sect[0].length):body[1];
  const pieces=rows.map((r,i)=>{let x=fillXml(content,re,r);if(i){x=x.replace(/(<wp:docPr\b[^>]*\bid=")(\d+)"/g,(m,a,id)=>a+(Number(id)+i*10000)+'"').replace(/(<w:bookmark(?:Start|End)\b[^>]*\bw:id=")(\d+)"/g,(m,a,id)=>a+(Number(id)+i*10000)+'"');}return x;});
  const pageBreak='<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  const merged=docXml.replace(body[0],'<w:body>'+pieces.join(pageBreak)+(sect?sect[0]:'')+'</w:body>');zip.file('word/document.xml',merged);
  const file=await freeName(outputDir,base+'-批量生成（'+rows.length+' 份）','.docx');await fs.writeFile(file,await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'}));
  let headerNote='';for(const n of otherParts){const f=new Set();fillXml(await zip.file(n).async('string'),re,null,f);if(f.size){headerNote='页眉 / 页脚里的占位符在合成文档里不会逐份替换；需要的话请选“每行一个文件”。';break;}}
  return {path:file,count:rows.length,used:[...found],note:headerNote};}
