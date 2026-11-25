// 2026-11-24 第二批新功能：检查与校对、号码校验、二维码 / 条形码、拼长图、批量重命名、表格数据工具、常用语、功能区分组。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import os from 'node:os';import path from 'node:path';import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
import {makeBook,makeApp} from './helpers/et-mock.mjs';
import {qrMatrix,code128Values,code128Modules,qrPng,barcodePng,makeCodes} from '../lib/codes.mjs';
import {decodePng,encodePng,stitchVertical} from '../lib/images.mjs';
import {planRename,applyRename,undoRename,newName} from '../lib/rename-files.mjs';

const read=f=>fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8');
function engine(file,name){const ctx={};ctx.window=ctx;vm.runInNewContext(read(file),ctx);return ctx[name];}
function bridge(a,host,confirm=()=>true,extra={}){const asked=[];const window={Application:a,confirm:m=>{asked.push(m);return confirm(m);},localStorage:{store:{},getItem(k){return this.store[k]??null;},setItem(k,v){this.store[k]=String(v);}}};vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/'+host+'/',search:''},URLSearchParams,Date,setTimeout,clearTimeout,...extra});return {T:window.OfficeTools,asked,window};}

test('错别字与标点：常见错别字、重复字、半角标点、括号不成对、半角片假名；正常写法不误报',()=>{
  const P=engine('proof-engine.js','OfficeProof');
  const t='我们以经按装好了,请大家确认.会议截止目前有（3人。。下周的的安排 ，见附件(说明)。。。ｱｲｳ';
  const kinds=P.check(t).map(x=>x.kind+':'+x.text);
  for(const k of ['typo:以经','typo:按装','punct:,','punct:.','typo:截止目前','bracket:（','punct:。。','repeat:的的','punct:(说明)','punct:。。。','width:ｱｲｳ'])assert.ok(kinds.includes(k),k);
  assert.equal(P.apply(t,P.check(t)),'我们已经安装好了，请大家确认。会议截至目前有（3人。下周的安排，见附件（说明）……アイウ');
  assert.equal(P.check('不了了之，是是非非。version 1.2.3 ok, e.g. test.').length,0);
  assert.equal(P.check('本日は会議です。資料は「別紙」のとおりです。').length,0,'日文正常句子不报');
  assert.equal(JSON.stringify(P.check('見ずらい資料をお送りしますがが、よろしく').map(x=>x.fix)),JSON.stringify(['見づらい','が']));
  assert.match(P.summary(P.check(t)),/错别字 3 处/);
});

test('号码校验：统一社会信用代码、身份证、法人番号 / T 番号、银行卡；自动判断种类',()=>{
  const I=engine('id-engine.js','OfficeIdCheck');
  assert.equal(I.check('91350100M000100Y43').ok,true);assert.match(I.check('91350100M000100Y44').msg,/校验码应为 3/);
  assert.equal(I.check('11010519491231002X').type,'id');assert.match(I.check('110105194912310021').msg,/应为 X/);
  assert.equal(I.check('7000012050002').ok,true);assert.equal(I.check('T7000012050002').type,'jpinvoice');assert.match(I.check('7000012050003').msg,/应为 6/);
  assert.equal(I.check('4111111111111111').ok,true);assert.equal(I.check('25312000000012345678').type,'invoice');
  assert.equal(I.check('９１３５ ０１００-Ｍ０００ １００Ｙ４３').ok,true,'全角、空格、横线也能认');assert.equal(I.check(''),null);assert.match(I.check('abc').msg,/认不出/);
});

test('二维码：版本随长度增加、三个定位图形、格式信息满足 BCH；条形码 Code 128 校验位正确、能还原',()=>{
  const m1=qrMatrix('A');assert.equal(m1.length,21);const m7=qrMatrix('x'.repeat(120));assert.equal((m7.length-17)/4,7);
  for(const m of [m1,m7,qrMatrix('办公工具箱 https://example.com/?q=1')]){const n=m.length;for(const [r,c] of [[0,0],[0,n-7],[n-7,0]]){assert.equal(m[r][c],true);assert.equal(m[r+1][c+1],false);assert.equal(m[r+3][c+3],true);}
    let f=0;for(let i=0;i<8;i++)f|=(m[8][n-1-i]?1:0)<<i;for(let i=8;i<15;i++)f|=(m[n-15+i][8]?1:0)<<i;f^=0x5412;let v=f;for(let i=14;i>=10;i--)if(v&(1<<i))v^=0x537<<(i-10);assert.equal(v,0,'格式信息校验');assert.equal(f>>13,0,'纠错等级 M');}
  assert.throws(()=>qrMatrix('x'.repeat(500)),/太长/);
  const vals=code128Values('INV2026-11-24');let sum=vals[0];for(let i=1;i<vals.length-2;i++)sum+=vals[i]*i;assert.equal(sum%103,vals.at(-2));assert.equal(vals.at(-1),106);
  assert.equal(code128Modules('x').length,46);assert.throws(()=>code128Values('中文'),/二维码/);
  const q=decodePng(qrPng('hello',4));assert.equal(q.width,(21+8)*4);const b=decodePng(barcodePng('ABC-123'));assert.equal(b.height,90);
});

test('批量二维码：每行一个 PNG，空行跳过，出错的记下来',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'codes-'));const r=await makeCodes({items:['https://a.example','','B-01'],kind:'qr',outputDir:dir});
  assert.equal(r.files.length,3);assert.equal(r.files[1],'');assert.ok(fs.existsSync(r.files[0])&&fs.existsSync(r.files[2]));
  const r2=await makeCodes({items:['OK-1','中文'],kind:'barcode',outputDir:dir});assert.equal(r2.errors.length,1);assert.equal(r2.errors[0].index,1);
});

test('拼长图：不同宽度按最宽的缩放（最宽 2000）、中间留白、不覆盖已有文件',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'stitch-'));const mk=(w,h,rgb,f)=>{const data=Buffer.alloc(w*h*4);for(let i=0;i<w*h;i++){data[i*4]=rgb[0];data[i*4+1]=rgb[1];data[i*4+2]=rgb[2];data[i*4+3]=255;}fs.writeFileSync(path.join(dir,f),encodePng({width:w,height:h,data}));return path.join(dir,f);};
  const a=mk(200,100,[255,0,0],'a.png'),b=mk(100,50,[0,0,255],'b.png'),out=path.join(dir,'长图.png');
  const r=await stitchVertical({paths:[a,b],output:out,gap:10});assert.equal(r.width,200);assert.equal(r.height,100+10+100);
  const img=decodePng(fs.readFileSync(r.path));assert.equal(img.data[0],255);const y=150;assert.equal(img.data[(y*200)*4+2],255,'第二张放大后是蓝色');assert.equal(img.data[(105*200)*4],255);assert.equal(img.data[(105*200)*4+1],255,'间隔是白色');
  const r2=await stitchVertical({paths:[a,b],output:out});assert.notEqual(r2.path,r.path);
});

test('批量重命名：预览、跳过冲突、互换名字不覆盖、记录后可以改回',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'rename-')),f=n=>path.join(dir,n);for(const n of ['初稿-合同.docx','初稿-报价.xlsx','已有.txt','合同.docx'])fs.writeFileSync(f(n),n);
  assert.equal(newName('C:/x/照片.jpg',{mode:'number',text:'旅行-#'},4,12),'旅行-05.jpg');assert.equal(newName('a.txt',{mode:'prefix',text:'2026-'},0),'2026-a.txt');
  const p=await planRename({paths:[f('初稿-合同.docx'),f('初稿-报价.xlsx')],rule:{mode:'replace',find:'初稿-',text:''}});
  assert.equal(p.count,1);assert.equal(p.plan[0].skip,'文件夹里已经有同名文件');
  const logDir=path.join(dir,'记录');const r=await applyRename({paths:[f('初稿-合同.docx'),f('初稿-报价.xlsx')],rule:{mode:'replace',find:'初稿-',text:''},logDir});
  assert.equal(r.done.length,1);assert.ok(fs.existsSync(f('报价.xlsx')));assert.ok(fs.existsSync(f('初稿-合同.docx')));assert.ok(fs.existsSync(r.log));
  const u=await undoRename({logDir});assert.equal(u.back.length,1);assert.ok(fs.existsSync(f('初稿-报价.xlsx')));
  // 互换：a→b、b→a
  fs.writeFileSync(f('1.txt'),'one');fs.writeFileSync(f('2.txt'),'two');const sw=await applyRename({paths:[f('1.txt'),f('2.txt')],rule:{mode:'number',start:'2',width:1},logDir});
  assert.equal(sw.done.length,2,'1→2 时 2 也在名单里，先改临时名再改新名，不会互相覆盖');assert.equal(fs.readFileSync(f('2.txt'),'utf8'),'one');assert.equal(fs.readFileSync(f('3.txt'),'utf8'),'two');
  await assert.rejects(planRename({paths:[f('已有.txt')],rule:{mode:'replace',find:''}}),/要替换的文字/);
});

test('表格：万元显示只改格式；号码校验写在右边空白列；15 位以上的数字提示已丢失',async()=>{
  const book=makeBook({S:[['金额','代码'],[123456,'91350100M000100Y43'],[9876543,'91350100M000100Y44'],[12,1234567890123456789]]}),a=makeApp(book,'A2:A4');
  const {T}=bridge(a,'et');assert.match(await T.run('wanUnit',{host:'et',unit:'wan'}),/3 个数字/);assert.equal(book.list[0].cell(2,1).NumberFormat,'0!.0,"万"');assert.equal(book.list[0].cell(2,1).Value2,123456);
  a.select('B2:B4');const msg=await T.run('idCheck',{host:'et',type:'auto'});assert.match(msg,/通过 1 个，有问题 2 个/);
  assert.match(book.list[0].cell(2,3).Value2,/^✓/);assert.match(book.list[0].cell(3,3).Value2,/校验码应为 3/);assert.match(book.list[0].cell(4,3).Value2,/15 位/);
  a.select('A2:A4');await assert.rejects(T.run('idCheck',{host:'et'}),/右边/,'右边有内容时不覆盖');
});

test('表格：考勤统计按记号计数（√ ✓ 算同一种），新表里有合计；原表不动',async()=>{
  const rows=[['姓名','1','2','3','4'],['张伟','√','✓','休','迟'],['李娜','√','假','√','√'],['张伟','','√','','']];
  const book=makeBook({S:rows}),a=makeApp(book,'A1:E4');const msg=await bridge(a,'et').T.run('attendanceStats',{host:'et'});
  const ws=book.list[1];assert.match(msg,/2 人、4 天/);assert.equal(ws.cell(1,1).Value2,'姓名');assert.equal(ws.cell(1,2).Value2,'√');
  const zw=[2,3,4,5,6].map(c=>ws.cell(2,c).Value2);assert.deepEqual(zw.slice(0,1),[3]);assert.equal(ws.cell(4,1).Value2,'合计');assert.equal(book.list[0].cell(2,3).Value2,'✓','原表不变');
});

test('表格：合并多个工作簿按表头对齐，加来源文件列，原文件只读打开后关闭',async()=>{
  const book=makeBook({S:[['x']]}),a=makeApp(book,'A1');const opened=[];
  const files={'C:/d/一月.xlsx':[['姓名','金额'],['张伟',10]],'C:/d/二月.xlsx':[['姓名','部门','金额'],['李娜','销售',20],['','','']]};
  a.FileDialog=()=>({Filters:{Clear(){},Add(){}},Show:()=>-1,SelectedItems:{Count:2,Item:i=>Object.keys(files)[i-1]}});
  a.Workbooks={Open:(f,u,ro)=>{opened.push([f,ro]);const wb=makeBook({Sheet1:files[f]});wb.Close=()=>{wb.closed=true;};return wb;}};
  const msg=await bridge(a,'et').T.run('mergeWorkbooks',{host:'et',mode:'stack'});
  assert.match(msg,/2 个文件的 2 行/);assert.ok(opened.every(x=>x[1]===true));const ws=book.list[1];
  assert.deepEqual([1,2,3,4].map(c=>ws.cell(1,c).Value2),['来源文件','姓名','金额','部门']);assert.deepEqual([1,2,3,4].map(c=>ws.cell(3,c).Value2),['二月','李娜',20,'销售']);
});

// 用段落和字符位置模拟 WPS 文字，供检查类工具使用。
function wordDoc(paras){
  const doc={text:paras.join('\r')+'\r',comments:[]};
  const range=(s,e)=>({Start:s,End:e,get Text(){return doc.text.slice(s,e);},set Text(v){doc.text=doc.text.slice(0,s)+v+doc.text.slice(e);},
    get Paragraphs(){const list=[];let at=0;for(const p of doc.text.split('\r').slice(0,-1)){const ps=at,pe=at+p.length+1;if(pe>s&&ps<Math.max(e,s+1))list.push({Range:range(ps,pe)});at=pe;}return {Count:list.length,Item:i=>list[i-1]};}});
  const d={Name:'a.docx',FullName:'C:\\x\\a.docx',get Content(){return range(0,doc.text.length);},Range:(s,e)=>range(s,e),
    Comments:{get Count(){return doc.comments.length;},Item:i=>{const c=doc.comments[i-1];return {Range:{Text:c.text},Delete(){doc.comments.splice(doc.comments.indexOf(c),1);}};},Add:(r,t)=>{doc.comments.push({at:r.Start,text:t});}},
    Revisions:{Count:0},BuiltInDocumentProperties:{Item:k=>({Value:k==='Author'?'山田':''})}};
  const a={ActiveDocument:d,Selection:{Text:'',Range:range(0,0)}};return {a,d,doc};
}
test('错别字与标点检查（文字）：加批注不改文字；直接改时只改能确定的；清除只删工具加的批注',async()=>{
  const {a,doc}=wordDoc(['我们以经按装好了,请确认。','（括号没关','正常的一段。']);doc.comments.push({at:0,text:'我自己的批注'});
  const {T}=bridge(a,'wps');let msg=await T.run('proofCheck',{host:'wps',mode:'comment'});
  assert.match(msg,/发现 4 处/);assert.equal(doc.comments.length,5);assert.ok(doc.comments.slice(1).every(c=>c.text.startsWith('【办公工具箱】')));assert.match(doc.text,/以经按装/,'没改文字');
  msg=await T.run('proofCheck',{host:'wps',mode:'clear'});assert.match(msg,/4 条/);assert.equal(doc.comments.length,1);
  msg=await T.run('proofCheck',{host:'wps',mode:'fix'});assert.match(msg,/已改正 3 处/);assert.match(doc.text,/我们已经安装好了，请确认。/);assert.equal(doc.comments.length,2,'括号不成对只加批注');
});
test('发送前检查：报告批注和作者；取消时不改；确认后清理',async()=>{
  const {a,d,doc}=wordDoc(['正文']);doc.comments.push({at:0,text:'意见'});d.Content.Find=null;
  const findless=()=>{throw Error('no find');};Object.defineProperty(d,'Content',{get(){return {Find:{ClearFormatting:findless},Text:doc.text};}});
  let props={Author:'山田'};d.BuiltInDocumentProperties={Item:k=>({get Value(){return props[k]||'';},set Value(v){props[k]=v;}})};
  let r=bridge(a,'wps',()=>false);let msg=await r.T.run('sendCheck',{host:'wps'});assert.match(msg,/批注 1 条/);assert.match(msg,/作者“山田”/);assert.match(msg,/没有做任何修改/);assert.equal(doc.comments.length,1);
  r=bridge(a,'wps');msg=await r.T.run('sendCheck',{host:'wps'});assert.match(msg,/删除批注/);assert.equal(doc.comments.length,0);assert.equal(props.Author,'');
});

test('常用语：内置按界面语言，自己写的会存下来排在最前，表格里不覆盖已有内容',async()=>{
  const book=makeBook({S:[['已有']]}),a=makeApp(book,'A1');let asked=0;const {T,window}=bridge(a,'et',()=>{asked++;return false;});
  const f=T.describe('insertPhrase','et').fields[0];assert.ok(f.options.some(o=>o.value==='特此通知。'));
  assert.equal(await T.run('insertPhrase',{host:'et',phrase:'特此通知。'}),'已取消。');assert.equal(asked,1);
  a.select('A2');await T.run('insertPhrase',{host:'et',custom:'请于周五前提交。'});assert.equal(book.list[0].cell(2,1).Value2,'请于周五前提交。');
  assert.equal(T.describe('insertPhrase','et').fields[0].options[0].value,'请于周五前提交。');assert.ok(window.localStorage.store['officeTools.phrases']);
});

test('功能区：常用工具除了各程序的常用按钮，只留打开工具箱、技巧导航、设置和工作室；文件、AI、检查与校对、数据工具、导出与瘦身各成一组；新按钮都有文字和图标',()=>{
  const xml=read('ribbon.xml'),text=JSON.parse(/window\.OfficeRibbonText=(\{.*\});?/s.exec(read('ribbon-text.js'))[1]);
  const g=xml.slice(xml.indexOf('<group id="toolsGroup"'),xml.indexOf('</group>',xml.indexOf('<group id="toolsGroup"')));
  assert.equal([...g.matchAll(/<button[^>]*id="([^"]+)"/g)].map(m=>m[1]).filter(id=>!/^c[wep]_/.test(id)).join(','),'toolsOpen,settingsOpen,studioOpen');
  for(const id of ['fileGroup','aiGroup','proofGroup','sheetDataGroup','slideOutGroup','sendCheck','proofCheck','compareDocs','exportComments','captionNumbers','fixListNumbering','idCheck','mergeWorkbooks','wanUnit','codesBatch','attendanceStats','slidesToImages','handoutPdf','slimDeck','batchRename','insertPhrase']){assert.ok(xml.includes('id="'+id+'"'),id);assert.ok(text[id]&&text[id].label,id);}
  const ctx={};ctx.window=ctx;vm.runInNewContext(read('tool-icons.js'),ctx);for(const id of ['sendCheck','codesBatch','slidesToImages','insertPhrase'])assert.ok(fs.existsSync(new URL('../public/'+ctx.ToolIcons.forAction(id),import.meta.url)),id);
  const main=read('main.js');assert.match(main,/id==='proofGroup'\)return h==='wps'/);assert.match(main,/id==='sheetDataGroup'\)return h==='et'/);assert.match(main,/notice\.html/);
});

test('Office 版：第二批工具登记进侧栏；万元显示、号码校验用 Office.js 写；做不到的给操作步骤',async()=>{
  const ctx={};ctx.window=ctx;vm.runInNewContext(['id-engine.js','proof-engine.js','office/office-docplus.js','sheet-logic.js','office/office-sheetplus.js','office/office-tools.js','office/office-batch2.js'].map(read).join('\n'),ctx);const T=ctx.OfficeAddinTools;
  for(const [h,ids] of Object.entries({excel:['idCheck','wanUnit','codesBatch','attendanceStats','mergeWorkbooks','insertPhrase'],word:['sendCheck','proofCheck','compareDocs','exportComments','captionNumbers','fixListNumbering','insertPhrase'],powerpoint:['insertPhrase','slidesToImages','handoutPdf','slimDeck']})){const list=T.list(h).map(x=>x.id);for(const id of ids)assert.ok(list.includes(id),h+' '+id);}
  // 模拟 Excel：选中一列
  const vals=[['91350100M000100Y43'],[''],['91350100M000100Y44']],right=[[''],[''],['']];let fmt=null,written=null;
  const range={values:vals,text:vals,rowCount:3,columnCount:1,load(){},set numberFormat(v){fmt=v;},getOffsetRange:()=>({getResizedRange:()=>({values:right,load(){},set values(v){written=v;},get values(){return right;},format:{}})})};
  const Excel={run:fn=>fn({workbook:{getSelectedRange:()=>range},sync:async()=>{}})};
  const msg=await T.run('excel','idCheck',{Excel},{type:'auto'});assert.match(msg,/通过 1 个，有问题 1 个/);assert.match(written[2][0],/校验码应为 3/);
  range.values=[[123456,'x']];range.columnCount=2;range.rowCount=1;assert.match(await T.run('excel','wanUnit',{Excel},{unit:'wan'}),/1 个数字/);assert.equal(fmt[0][0],'0"."0,"万"'); await T.run('excel','wanUnit',{Excel},{unit:'yi'});assert.equal(fmt[0][0],'0"."00,,"亿"');
  assert.match(await T.run('word','compareDocs',{},{}),/审阅 → 比较/);
});

test('号码自动识别：错误出生日期仍按身份证报错，不回退信用代码误报通过',()=>{
  const I=engine('id-engine.js','OfficeIdCheck');
  for(const value of ['11010519491331002X','11010519490231002X','11010520991231002X']) {
    const r=I.check(value,'auto'); assert.equal(r.type,'id'); assert.equal(r.ok,false); assert.match(r.msg,/出生日期/);
    assert.match(I.check(value,'auto',{lang:'ja'}).msg,/生年月日/);
  }
  assert.equal(I.check('91350100M000100Y43').ok,true);
});
test('Office 二维码和条形码同时设置宽高、保持比例且不超出目标单元格',async()=>{
  const ctx={};ctx.window=ctx;vm.runInNewContext(['office/office-tools.js','office/office-batch2.js'].map(read).join('\n'),ctx);
  for(const kind of ['qr','barcode']) {
    const shapes=[],cell={left:30,top:40,width:64,height:70,load(){}};
    const out={values:[['']],format:{},load(){},getCell(){return cell;}};
    const range={values:[['QA-001']],text:[['QA-001']],rowCount:1,columnCount:1,load(){},getOffsetRange(){return {getResizedRange(){return out;}};}};
    const sh={addImage(){const img={width:kind==='qr'?240:400,height:kind==='qr'?240:100,load(){}};shapes.push(img);return img;}};
    const Excel={run:fn=>fn({workbook:{getSelectedRange:()=>range,worksheets:{getActiveWorksheet:()=>({shapes:sh})}},sync:async()=>{}})};
    const fetch=async url=>({ok:true,json:async()=>url==='/api/session'?{token:'test'}:{data:['image'],folder:'qa',errors:[{index:3,error:'条形码只能用英文字母、数字和常用符号；中文请用二维码。'}]}});
    const message=await ctx.OfficeAddinTools.run('excel','codesBatch',{Excel},{kind,confirmed:true,fetch});assert.match(message,/第 4 行/);assert.match(message,/中文请用二维码/);
    assert.equal(shapes.length,1);const img=shapes[0];assert.equal(img.width/img.height,kind==='qr'?1:4);
    assert.ok(img.width<=58&&img.height<=64);assert.equal(img.top,43);assert.equal(img.placement,'TwoCell');
  }
});

test('Office 常用语：确认前不保存；成功后置顶、重载持久化、清空保留内置',async()=>{
  const store={};const localStorage={getItem:k=>store[k],setItem:(k,v)=>store[k]=v};
  function load(){const ctx={localStorage};ctx.window=ctx;vm.runInNewContext(['office/office-tools.js','office/office-batch2.js'].map(read).join('\n'),ctx);return ctx.OfficeAddinTools;}
  let T=load();const cell={values:[['原内容']],load(){}};
  const Excel={run:fn=>fn({workbook:{getActiveCell:()=>cell},sync:async()=>{}})};
  const options=()=>T.list('excel').find(x=>x.id==='insertPhrase').fields[0].options;
  assert.ok((await T.run('excel','insertPhrase',{Excel},{custom:'测试常用语'})).confirm);
  assert.equal(store['officeTools.phrases'],undefined);assert.equal(cell.values[0][0],'原内容');
  await T.run('excel','insertPhrase',{Excel},{custom:'测试常用语',confirmed:true});assert.equal(cell.values[0][0],'测试常用语');
  T=load();assert.equal(options()[0].label,'★ 测试常用语');
  await T.run('excel','insertPhrase',{Excel},{custom:'测试常用语',confirmed:true});assert.equal(options().filter(x=>x.value==='测试常用语').length,1);
  await T.run('excel','insertPhrase',{}, {phrase:'__clear'});assert.ok(!options().some(x=>x.label.startsWith('★')));assert.ok(options().some(x=>x.value==='特此通知。'));assert.equal(cell.values[0][0],'测试常用语');
});

test('题注：图片题注不能代替表题注，表题注不能代替图片题注',()=>{
  for(const [kind,label,expected] of [['table','图',1],['table','図',1],['table','表',0],['pic','表',1],['pic','图',0],['pic','図',0]]) {
    let tool,asked='';
    const root={OfficeTools:{lib:{ask:m=>{asked=m;return false;}},register(t){if(t.id==='captionNumbers')tool=t;}}};
    vm.runInNewContext(read('tools/proof.js'),{window:root});
    const adjacent={Range:{Text:label+' 1 测试\r',Start:20,End:30}};
    const picPara={Range:{Start:10,End:20}};
    const d={Content:{Text:'测试'},Paragraphs:{Count:0},Tables:{Count:kind==='table'?1:0,Item:()=>({NestingLevel:1,Range:{Start:31}})},InlineShapes:{Count:kind==='pic'?1:0,Item:()=>({Type:3,Range:{Information:()=>false,Paragraphs:{Item:()=>picPara}}})},Range:()=>({Paragraphs:{Item:()=>adjacent}})};
    tool.hosts.wps({a:{},d,options:{}});
    assert.equal(Boolean(asked),Boolean(expected),kind+' adjacent '+label);
    if(expected)assert.match(asked,kind==='table'?/1 个表格/:/1 张图片/);
  }
});

test('题注：插入段落会扩展 Range.End，文字仍须写入新段落而不是后方表格',()=>{
  let tool,insertAt=-1,fieldAt=-1;
  const root={OfficeTools:{lib:{ask:()=>true},register(t){if(t.id==='captionNumbers')tool=t;}}};
  vm.runInNewContext(read('tools/proof.js'),{window:root});
  const previous={Range:{Text:'普通段落\r',Start:10,End:20,InsertParagraphAfter(){this.End=21;}}};
  const d={Content:{Text:'测试'},Paragraphs:{Count:0},Tables:{Count:1,Item:()=>({NestingLevel:1,Range:{Start:20}})},InlineShapes:{Count:0},Fields:{Add(r){fieldAt=r.s;},Update(){}},Range(s){return {s,Paragraphs:{Item:()=>previous},InsertBefore(v){assert.equal(s,19);assert.equal(v,'\r');previous.Range.End=21;},InsertAfter(v){insertAt=s;assert.equal(v,'表 ');}};}};
  const result=tool.hosts.wps({a:{},d,options:{}});
  assert.equal(insertAt,20);assert.equal(fieldAt,22);assert.match(result,/1 个表题/);
});

test('发送前清理：接受修订和删除隐藏文字后重新定位，不能用旧位置清高亮',async()=>{
  let hidden=[8,12],highlight=[12,16],cleared=false;
  const {a,d}=wordDoc(['0123456789abcdefghijkl']);
  d.TrackRevisions=true;
  d.Revisions={Count:1,AcceptAll(){hidden=hidden.map(x=>x-3);highlight=highlight.map(x=>x-3);this.Count=0;}};
  Object.defineProperty(d,'Content',{get(){
    const r={Start:0,End:24,Collapse(){this.Start=this.End;}};
    r.Find={Font:{},ClearFormatting(){},Execute(){const target=this.Highlight?highlight:hidden;if(!target||target[0]<r.Start)return false;[r.Start,r.End]=target;return true;}};
    return r;
  }});
  d.Range=(s,e)=>({Delete(){assert.deepEqual([s,e],hidden);const len=e-s;hidden=null;highlight=highlight.map(x=>x-len);},set HighlightColorIndex(v){assert.equal(v,0);assert.deepEqual([s,e],highlight);cleared=true;}});
  await bridge(a,'wps').T.run('sendCheck',{host:'wps'});
  assert.equal(hidden,null);assert.equal(cleared,true);assert.equal(d.TrackRevisions,false);
});
