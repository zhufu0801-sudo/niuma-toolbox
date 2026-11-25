// 常用文书模板（2026-11-21）：中文 请假单、报销单、会议纪要、周报；日文 議事録、稟議書、発注書、日報、始末書。
// 运行：node scripts/templates/build-forms.mjs  → public/templates/forms/*.docx|xlsx 与 index.json
// 内容是通用格式（空栏和示例提示），不含任何真实公司或个人信息。
import {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,AlignmentType,VerticalAlign,HeightRule,BorderStyle} from 'docx';
import ExcelJS from 'exceljs';import fs from 'node:fs/promises';import path from 'node:path';import {fileURLToPath} from 'node:url';
import {rmbUpperFormula} from '../../lib/business.mjs';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'../..'),out=path.join(root,'public/templates/forms');await fs.mkdir(out,{recursive:true});

const FONT={zh:{body:'宋体',head:'黑体'},ja:{body:'Yu Mincho',head:'Yu Gothic'}};
let L='zh';
const run=(text,o={})=>new TextRun({text,font:{ascii:o.head?FONT[L].head:FONT[L].body,eastAsia:o.head?FONT[L].head:FONT[L].body,hAnsi:o.head?FONT[L].head:FONT[L].body},size:(o.size||10.5)*2,bold:!!o.bold,color:o.color});
const para=(text,o={})=>new Paragraph({alignment:o.align==='center'?AlignmentType.CENTER:o.align==='right'?AlignmentType.RIGHT:AlignmentType.LEFT,spacing:{before:o.before??0,after:o.after??80,line:o.line??300},children:(Array.isArray(text)?text:[text]).map(t=>typeof t==='string'?run(t,o):t)});
const title=t=>para(t,{head:true,size:20,bold:true,align:'center',after:240});
const heading=t=>para(t,{head:true,size:12,bold:true,before:200,after:100});
const note=t=>para(t,{size:9,color:'5D6878',after:120});
const W=9000; // A4 可用宽度约 9000 dxa（左右 2.2 cm 页边距）
function cell(text,{w,bold=false,shade=false,align='left',h,span=1,rowSpan}={}){return new TableCell({width:w?{size:w,type:WidthType.DXA}:undefined,columnSpan:span>1?span:undefined,rowSpan,verticalAlign:VerticalAlign.CENTER,shading:shade?{fill:'EEF2F7'}:undefined,margins:{top:60,bottom:60,left:100,right:100},children:String(text).split('\n').map(line=>para(line,{bold,align,after:0,line:276}))});}
function table(rows,{widths,heights=[]}={}){const total=widths.reduce((a,b)=>a+b,0);return new Table({width:{size:total,type:WidthType.DXA},columnWidths:widths,rows:rows.map((r,i)=>new TableRow({height:heights[i]?{value:heights[i],rule:HeightRule.ATLEAST}:undefined,children:r.map(c=>c instanceof TableCell?c:cell(c))}))});}
// 表单行：标签（灰底）+ 填写栏，交替。
function form(pairs,cols=4){const lw=1400,vw=Math.floor((W-lw*(cols/2))/(cols/2));return table(pairs.map(row=>row.flatMap(([k,v,span])=>[cell(k,{w:lw,shade:true,bold:true}),cell(v||'',{w:vw,span:span||1})])),{widths:Array.from({length:cols},(_,i)=>i%2?vw:lw)});}
function grid(header,rows,widths,height=420){return table([header.map((h,i)=>cell(h,{w:widths[i],shade:true,bold:true,align:'center'})),...Array.from({length:rows},(_,r)=>header.map((h,i)=>cell(i===0&&/序号|No/.test(header[0])?String(r+1):'',{w:widths[i],align:i===0?'center':'left'})))],{widths,heights:[0,...Array(rows).fill(height)]});}
function bigBox(label,height=1600,hint=''){return table([[cell(label,{w:1400,shade:true,bold:true}),cell(hint,{w:W-1400})]],{widths:[1400,W-1400],heights:[height]});}
// 押印欄：靠右，职位一行 + 盖章格一行。
function stampBox(labels){const w=1134;return new Table({width:{size:w*labels.length,type:WidthType.DXA},columnWidths:labels.map(()=>w),alignment:AlignmentType.RIGHT,rows:[new TableRow({children:labels.map(l=>cell(l,{w,align:'center',shade:true}))}),new TableRow({height:{value:1134,rule:HeightRule.EXACT},children:labels.map(()=>cell('',{w}))})]});}
const gap=()=>para('',{after:60});
function doc(children){return new Document({styles:{default:{document:{run:{font:FONT[L].body,size:21}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1247,right:1247}}},children}]});}

const docs={
 'leave-zh':()=>[title('请假申请单'),form([[['姓名'],['部门']],[['职位'],['联系电话']],[['申请日期','    年    月    日',3]]]),gap(),
   bigBox('请假类型',600,'□ 事假　□ 病假　□ 年假　□ 调休　□ 婚假　□ 产假 / 陪产假　□ 丧假　□ 其他：________'),
   bigBox('请假时间',700,'自　　　年　　月　　日　　时　起\n至　　　年　　月　　日　　时　止　　共　　天　　小时'),
   bigBox('请假事由',1600),bigBox('工作交接',900,'交接人：____________　　交接内容：'),gap(),
   table([[cell('部门负责人意见',{w:3000,shade:true,bold:true,align:'center'}),cell('人事部门意见',{w:3000,shade:true,bold:true,align:'center'}),cell('总经理意见',{w:3000,shade:true,bold:true,align:'center'})],[cell('\n\n签字：\n日期：',{w:3000}),cell('\n\n签字：\n日期：',{w:3000}),cell('\n\n签字：\n日期：',{w:3000})]],{widths:[3000,3000,3000]}),
   gap(),para('申请人签字：______________　　日期：______________',{align:'right'}),note('说明：病假请附医院证明；请假 3 天以上需总经理审批。请按公司规定提前提交。')],
 'minutes-zh':()=>[title('会议纪要'),form([[['会议名称','',3]],[['会议时间'],['会议地点']],[['主持人'],['记录人']],[['参会人员','',3]],[['缺席人员','',3]]]),
   heading('一、会议议题'),para('1. '),para('2. '),heading('二、讨论情况'),para('（按议题记录主要观点和讨论结果）'),para(''),heading('三、会议决议'),para('1. '),para('2. '),
   heading('四、待办事项'),grid(['序号','事项','负责人','完成时间','状态'],5,[700,4300,1400,1500,1100]),heading('五、下次会议'),para('时间：　　　　　地点：　　　　　议题：'),
   gap(),para('记录人：__________　　审核：__________',{align:'right'})],
 'weekly-zh':()=>[title('工作周报'),form([[['姓名'],['部门']],[['周期','    年    月    日 至     月    日',3]]]),
   heading('一、本周完成'),grid(['序号','工作内容','完成情况','备注'],5,[700,4600,2000,1700]),
   heading('二、问题与需要的支持'),bigBox('问题 / 风险',1100),bigBox('需要的支持',900),
   heading('三、下周计划'),grid(['序号','计划内容','预计完成时间','备注'],4,[700,4600,2000,1700]),heading('四、其他'),para('')],
 'minutes-ja':()=>[stampBox(['作成','確認','承認']),title('議事録'),form([[['会議名','',3]],[['日時','令和　年　月　日（　）　：　～　：'],['場所']],[['出席者','',3]],[['欠席者'],['記録者']]]),
   heading('1. 議題'),para('（1）'),para('（2）'),heading('2. 決定事項'),para('・'),para('・'),heading('3. 確認事項・意見'),para('・'),
   heading('4. 宿題（ToDo）'),grid(['No.','内容','担当','期限','状況'],5,[700,4300,1400,1500,1100]),heading('5. 次回予定'),para('日時：　　　　　場所：　　　　　議題：'),para('以上',{align:'right'})],
 'ringi-ja':()=>[stampBox(['社長','部長','課長','担当']),title('稟　議　書'),form([[['稟議No.'],['起案日','令和　年　月　日']],[['所属'],['起案者','　　　　　　　印']],[['決裁期限','令和　年　月　日',3]]]),gap(),
   bigBox('件　名',600),bigBox('稟議内容',1800,'（何を、いつ、どのように行うか）'),bigBox('目的・理由',1400,'（必要性、期待される効果）'),
   form([[['金額（税込）','　　　　　　　　円'],['予算','□ 予算内　□ 予算外']],[['支払先・取引先'],['実施時期']]]),bigBox('添付資料',600,'□ 見積書　□ カタログ　□ 契約書案　□ その他（　　　）'),bigBox('備考',800),
   gap(),table([[cell('決裁',{w:1400,shade:true,bold:true,align:'center'}),cell('□ 承認　□ 条件付承認　□ 否認　□ 保留\n条件・コメント：',{w:W-1400})]],{widths:[1400,W-1400],heights:[900]})],
 'daily-ja':()=>[stampBox(['上長確認']),title('業務日報'),form([[['日付','令和　年　月　日（　）'],['氏名']],[['所属','',3]]]),
   heading('本日の業務'),grid(['時間','業務内容','結果・進捗'],7,[1500,4500,3000],380),heading('成果・進捗'),bigBox('成果',900),
   heading('課題・問題点'),bigBox('課題',900,'（発生した問題、対応状況、相談したいこと）'),heading('明日の予定'),grid(['時間','予定','備考'],4,[1500,4500,3000],380),heading('所感'),bigBox('所感',700)],
 'apology-ja':()=>[para('令和　　年　　月　　日',{align:'right'}),para('株式会社〇〇〇〇',{after:0}),para('代表取締役　〇〇　〇〇　殿',{after:240}),
   para('所属：____________________',{align:'right',after:0}),para('氏名：____________________　印',{align:'right',after:300}),title('始　末　書'),
   para('このたび、私は令和　　年　　月　　日、下記のとおり（不始末の内容）を起こし、会社に多大なご迷惑をおかけしましたことを深くお詫び申し上げます。',{line:360}),para('記',{align:'center',before:120,after:120}),
   form([[['発生日時','令和　年　月　日（　）　：　頃',3]],[['発生場所','',3]]]),bigBox('経　緯',1600,'（いつ、どこで、何が起きたか、事実を順に）'),bigBox('原　因',1200),bigBox('今後の対策',1400,'（同じことを繰り返さないために行うこと）'),
   para('今後はこのようなことを二度と起こさぬよう十分注意いたしますので、何卒寛大なご処置をお願い申し上げます。',{before:200,line:360}),para('以上',{align:'right'}),
   note('※ 事実を正確に書き、言い訳にならないよう簡潔にまとめます。会社の規定様式がある場合はそちらを使ってください。')]
};
const index=[];
const META={'leave-zh':['请假单','zh'],'minutes-zh':['会议纪要','zh'],'weekly-zh':['工作周报','zh'],'minutes-ja':['議事録','ja'],'ringi-ja':['稟議書','ja'],'daily-ja':['業務日報','ja'],'apology-ja':['始末書','ja']};
for(const [id,make] of Object.entries(docs)){const [t,lang]=META[id];L=lang;const buf=await Packer.toBuffer(doc(make()));await fs.writeFile(path.join(out,id+'.docx'),buf);index.push({id,title:t,lang,file:id+'.docx',kind:'forms'});}

// ---- 表格：报销单、発注書（带公式） ----
async function book(id,t,lang,build){const wb=new ExcelJS.Workbook();wb.creator='Office Toolbox';wb.calcProperties.fullCalcOnLoad=true;const ws=wb.addWorksheet(t,{pageSetup:{paperSize:9,orientation:'portrait',fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:0.6,right:0.6,top:0.7,bottom:0.7,header:0.3,footer:0.3}}});
  const font=lang==='ja'?'Yu Gothic':'Microsoft YaHei';ws.properties.defaultRowHeight=20;build(ws,font);ws.eachRow(r=>r.eachCell(c=>{c.font={name:font,size:10.5,...(c.font||{}),name:font};}));
  await wb.xlsx.writeFile(path.join(out,id+'.xlsx'));index.push({id,title:t,lang,file:id+'.xlsx',kind:'forms'});}
const thin={style:'thin',color:{argb:'FF8A94A3'}},box={top:thin,left:thin,bottom:thin,right:thin};
function frame(ws,range){const [a,b]=range.split(':'),c1=ws.getCell(a),c2=ws.getCell(b);for(let r=c1.row;r<=c2.row;r++)for(let c=c1.col;c<=c2.col;c++)ws.getCell(r,c).border=box;}
function m(ws,range,v,o={}){ws.mergeCells(range);const c=ws.getCell(range.split(':')[0]);c.value=v;if(o.bold||o.size)c.font={bold:!!o.bold,size:o.size||10.5};c.alignment={vertical:'middle',horizontal:o.align||'left',wrapText:true};if(o.fill)c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEEF2F7'}};}
await book('expense-zh','费用报销单','zh',ws=>{ws.columns=[{width:7},{width:13},{width:14},{width:30},{width:14},{width:12}];
  m(ws,'A1:F2','费用报销单',{bold:true,size:20,align:'center'});m(ws,'A3:F3','填写后打印签字；金额合计和大写会自动计算。');
  m(ws,'A4:A4','报销人',{fill:1});m(ws,'B4:C4','');m(ws,'D4:D4','部门',{fill:1,align:'right'});m(ws,'E4:F4','');m(ws,'A5:A5','日期',{fill:1});m(ws,'B5:C5','');m(ws,'D5:D5','附单据（张）',{fill:1,align:'right'});m(ws,'E5:F5','');frame(ws,'A4:F5');
  const head=['序号','日期','费用类别','摘要（用途）','金额（元）','备注'];head.forEach((h,i)=>{const c=ws.getCell(7,i+1);c.value=h;c.font={bold:true};c.alignment={horizontal:'center',vertical:'middle'};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEEF2F7'}};});
  for(let r=8;r<=17;r++){ws.getCell(r,1).value=r-7;ws.getCell(r,1).alignment={horizontal:'center'};ws.getCell(r,2).numFmt='yyyy/m/d';ws.getCell(r,5).numFmt='#,##0.00';}
  ws.getCell('C8').value='［差旅 / 交通 / 餐费 / 办公用品…］';ws.getCell('C8').font={color:{argb:'FF8A94A3'}};frame(ws,'A7:F17');
  m(ws,'A18:D18','合计（元）',{bold:true,align:'right'});ws.getCell('E18').value={formula:'SUM(E8:E17)',result:0};ws.getCell('E18').numFmt='#,##0.00';ws.getCell('E18').font={bold:true};
  m(ws,'A19:B19','合计大写',{fill:1});m(ws,'C19:F19',{formula:rmbUpperFormula('E18'),result:'零元整'});frame(ws,'A18:F19');
  m(ws,'A21:B21','部门负责人',{fill:1,align:'center'});m(ws,'A22:B23','');m(ws,'C21:C21','财务审核',{fill:1,align:'center'});m(ws,'C22:C23','');m(ws,'D21:D21','总经理',{fill:1,align:'center'});m(ws,'D22:D23','');m(ws,'E21:F21','出纳',{fill:1,align:'center'});m(ws,'E22:F23','');frame(ws,'A21:F23');
  m(ws,'A25:F25','领款人签字：______________　　　日期：______________',{align:'right'});ws.getRow(1).height=22;ws.getRow(22).height=30;ws.getRow(23).height=30;ws.pageSetup.printArea='A1:F25';});
await book('order-ja','発注書','ja',ws=>{ws.columns=[{width:10},{width:30},{width:8},{width:7},{width:13},{width:15}];
  m(ws,'A1:F2','発　注　書',{bold:true,size:20,align:'center'});m(ws,'A3:F3','入力例です。発注前に品目・数量・単価・納期を確認してください。');
  m(ws,'A5:C5','［取引先名］　御中',{size:13});m(ws,'D5:F5','発注日：令和　年　月　日',{align:'right'});m(ws,'D6:F6','発注番号：',{align:'right'});m(ws,'D7:F7','［会社名］',{align:'right'});m(ws,'D8:F8','担当：　　　　　　印',{align:'right'});
  m(ws,'A9:C9','下記のとおり発注いたします。');
  [['納期','令和　年　月　日'],['納品場所',''],['支払条件','月末締め翌月末払い']].forEach((x,i)=>{m(ws,'A'+(10+i)+':A'+(10+i),x[0],{fill:1});m(ws,'B'+(10+i)+':C'+(10+i),x[1]);});frame(ws,'A10:C12');
  const head=['No.','品目','数量','単位','単価（税抜）','金額（税抜）'];head.forEach((h,i)=>{const c=ws.getCell(14,i+1);c.value=h;c.font={bold:true};c.alignment={horizontal:'center'};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFEEF2F7'}};});
  for(let r=15;r<=24;r++){ws.getCell(r,1).value=r-14;ws.getCell(r,1).alignment={horizontal:'center'};ws.getCell(r,6).value={formula:'IF(OR(C'+r+'="",E'+r+'=""),"",C'+r+'*E'+r+')',result:''};ws.getCell(r,5).numFmt='#,##0';ws.getCell(r,6).numFmt='#,##0';}
  ws.getCell('B15').value='［品目を入力（例：コピー用紙 A4）］';ws.getCell('B15').font={color:{argb:'FF8A94A3'}};frame(ws,'A14:F24');
  [['小計','SUM(F15:F24)'],['消費税（10%）','ROUNDDOWN(F25*0.1,0)'],['合計','F25+F26']].forEach((x,i)=>{const r=25+i;m(ws,'D'+r+':E'+r,x[0],{fill:1,align:'right',bold:i===2});ws.getCell('F'+r).value={formula:x[1],result:0};ws.getCell('F'+r).numFmt='#,##0';if(i===2)ws.getCell('F'+r).font={bold:true};});frame(ws,'D25:F27');
  m(ws,'A29:A29','備考',{fill:1});m(ws,'B29:F31','');frame(ws,'A29:F31');ws.getRow(1).height=22;ws.pageSetup.printArea='A1:F31';});

index.sort((a,b)=>a.lang.localeCompare(b.lang)||a.id.localeCompare(b.id));
await fs.writeFile(path.join(out,'index.json'),JSON.stringify(index,null,2));console.log('文书模板：'+index.length+' 套');
