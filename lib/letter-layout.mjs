// 日文竖排书信（ご案内・挨拶状）的排版规则。竖排里“左右对齐”变成“上下对齐”：END = 靠下。
// 规则：结语（敬具・以上 等）靠下；“記”居中；结语后面的日期往下缩两格；
// 日期后面的署名靠下；以“様・御中・殿”结尾的宛名靠上、稍大。数字和字母换成全角，避免横倒。
import {Paragraph,TextRun,AlignmentType} from 'docx';

const CLOSING=/^(敬具|敬白|拝具|草々|早々|かしこ|以上|謹言|謹白)$/;
const DATE=/^(令和|平成|昭和)?[〇一二三四五六七八九十元\d０-９]+年[〇一二三四五六七八九十\d０-９]+月([〇一二三四五六七八九十\d０-９]+日|吉日|上旬|中旬|下旬)?$/;
const ADDRESSEE=/(様|御中|殿)$/;
export const fullWidth=s=>String(s).replace(/[0-9A-Za-z]/g,c=>String.fromCharCode(c.charCodeAt(0)+0xFEE0)).replace(/-/g,'－');

export function letterLines(text){
  const lines=String(text||'').replace(/\r\n?/g,'\n').split('\n').map(s=>s.replace(/\s+$/,''));
  const out=[];let after=false,dated=false;
  for(const raw of lines){const t=raw.trim();let role='body';
    if(!t)role='blank';
    else if(t==='記')role='ki';
    else if(CLOSING.test(t)){role='closing';after=true;}
    else if(DATE.test(t)){role='date';dated=true;}
    else if(ADDRESSEE.test(t)&&t.length<=30&&(after||dated))role='addressee';
    else if(dated&&!/^記/.test(t))role='sender';
    if(role==='ki'){dated=false;}
    out.push({text:raw,role});}
  return out;
}

export function letterParagraphs(text,{title=''}={}){
  const P=[];if(title)P.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:200},children:[new TextRun({text:fullWidth(title),bold:true,size:32})]}));
  for(const l of letterLines(text)){
    const t=fullWidth(l.text);
    if(l.role==='blank'){P.push(new Paragraph(''));continue;}
    const base={spacing:{after:0,line:360}};
    if(l.role==='closing')P.push(new Paragraph({...base,alignment:AlignmentType.END,children:[new TextRun(t.trim())]}));
    else if(l.role==='ki')P.push(new Paragraph({...base,alignment:AlignmentType.CENTER,children:[new TextRun(t.trim())]}));
    else if(l.role==='date')P.push(new Paragraph({...base,indent:{start:480},children:[new TextRun(t.trim())]}));
    else if(l.role==='sender')P.push(new Paragraph({...base,alignment:AlignmentType.END,children:[new TextRun(t.trim())]}));
    else if(l.role==='addressee')P.push(new Paragraph({...base,spacing:{before:240,line:360},children:[new TextRun({text:t.trim(),size:28})]}));
    else P.push(new Paragraph({...base,children:[new TextRun(t)]}));
  }
  return P;
}

// 和暦の日付（漢数字）。例：令和八年十一月十九日
const KD='〇一二三四五六七八九';
export function kanjiNumber(n){n=Number(n);if(n<10)return KD[n];if(n<20)return '十'+(n%10?KD[n%10]:'');if(n<100)return KD[Math.floor(n/10)]+'十'+(n%10?KD[n%10]:'');return String(n).split('').map(d=>KD[d]).join('');}
export function wareki(date=new Date(),{day=true,kichijitsu=false}={}){
  const y=date.getFullYear()-2018,ys=y===1?'元':kanjiNumber(y);
  return '令和'+ys+'年'+kanjiNumber(date.getMonth()+1)+'月'+(kichijitsu?'吉日':day?kanjiNumber(date.getDate())+'日':'');
}
export const JIKOU=['厳寒','余寒','早春','陽春','新緑','梅雨','盛夏','残暑','初秋','秋冷','晩秋','師走'];
