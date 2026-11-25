// 找一个能写中文 / 日文的系统字体，给生成 PDF 用。Windows、macOS、Linux 都有候选；可以用 COACH_PDF_FONT 指定（.ttc 可写成 “文件路径#字体名”）。
import fs from 'node:fs';
import path from 'node:path';
import PDFDocument from 'pdfkit';
export function fontCandidates(env=process.env,platform=process.platform){
  const list=[];
  if(env.COACH_PDF_FONT){const [file,family]=String(env.COACH_PDF_FONT).split('#');list.push({file,family:family||undefined});}
  const win=path.win32.join(env.WINDIR||env.SystemRoot||'C:\\Windows','Fonts');
  const w=(f,family)=>list.push({file:platform==='win32'?path.win32.join(win,f):path.join(win,f),family});
  if(platform==='win32'){w('msyh.ttc','MicrosoftYaHei');w('simhei.ttf');w('simsun.ttc','SimSun');w('YuGothM.ttc','YuGothic-Medium');w('msgothic.ttc','MS-Gothic');w('meiryo.ttc','Meiryo');}
  if(platform==='darwin'){for(const [f,family] of [['/System/Library/Fonts/PingFang.ttc','PingFangSC-Regular'],['/System/Library/Fonts/Hiragino Sans GB.ttc','HiraginoSansGB-W3'],['/System/Library/Fonts/STHeiti Light.ttc','STHeitiSC-Light'],['/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc','HiraginoSans-W3'],['/Library/Fonts/Arial Unicode.ttf',undefined],['/System/Library/Fonts/Supplemental/Arial Unicode.ttf',undefined]])list.push({file:f,family});}
  for(const [f,family] of [['/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc','NotoSansCJKsc-Regular'],['/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc','NotoSansCJKsc-Regular'],['/usr/share/fonts/truetype/wqy/wqy-microhei.ttc','WenQuanYiMicroHei']])list.push({file:f,family});
  return list;
}
let cached=null;
// 返回 {file,family}；都不能用时返回 null。逐个试着让 pdfkit 打开，字体名不对的 .ttc 会被跳过。
export function findCjkFont(env=process.env,platform=process.platform){
  if(cached&&!env.COACH_PDF_FONT)return cached;
  for(const c of fontCandidates(env,platform)){if(!c.file||!fs.existsSync(c.file))continue;try{const doc=new PDFDocument();doc.font(c.file,c.family);doc.end();if(!env.COACH_PDF_FONT)cached=c;return c;}catch{}}
  return null;
}
