// 2026-11-20 客户体验测试（李老板、扫描件用户、职场小白）发现的问题，修好后防止回退。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import os from 'node:os';import path from 'node:path';import vm from 'node:vm';
import ExcelJS from 'exceljs';
import {toolsCode} from './helpers/tools-code.mjs';
import {openDocument,wpsProgramFor} from '../lib/open-document.mjs';
import {outputDirectory,validateSettings,defaultSettings} from '../lib/settings.mjs';

function bridge(a,host,confirm=()=>true){const asked=[];const window={Application:a,confirm:m=>{asked.push(m);return confirm(m);}};vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/'+host+'/',search:''},URLSearchParams,Date,setTimeout,clearTimeout});return {T:window.OfficeTools,asked};}

function sheet(values,extra={}){
  const writes={};const range={Rows:{Count:values.length},Columns:{Count:values[0].length},MergeCells:false,HasFormula:false,Address:'A1:D9',Value2:values,
    Cells:{Item:(r,c)=>({Resize:(rows,cols)=>({set Value2(v){writes.top={r,rows,cols,v};},ClearContents(){writes.clear={r,rows,cols};}})})},...extra};
  return {a:{ActiveWorkbook:{Name:'s.xlsx',FullName:'C:/x/s.xlsx'},ActiveSheet:{Name:'S'},Selection:range},writes,range};
}

test('删除重复行：电话、以 0 开头的编号、日期写回时仍是文字，不会变成 1.38E+10',async()=>{
  const {a,writes}=sheet([['姓名','电话','编号','日期'],['张伟','13800000001','0123','2026/11/20'],['李娜','021-12345678','0456','11:30'],['张伟','13800000001','0123','2026/11/20']]);
  const msg=await bridge(a,'et').T.run('removeDuplicates',{host:'et'});
  assert.equal(JSON.stringify(writes.top.v),JSON.stringify([['姓名','电话','编号','日期'],['张伟',"'13800000001","'0123","'2026/11/20"],['李娜',"'021-12345678","'0456","'11:30"]]));
  assert.match(msg,/删除 1 行/);
});
test('删除重复行：只差空格的行不删但会提示；从界面执行时只确认一次',async()=>{
  const {a}=sheet([['姓名'],['李娜'],[' 李娜'],['王强'],['王强']]);
  const {T,asked}=bridge(a,'et');const msg=await T.runFromUi('removeDuplicates',{host:'et'});
  assert.equal(asked.length,1,'只弹工具自己的确认');assert.match(asked[0],/1 行只是多了空格/);assert.match(msg,/1 行只差空格/);
});
test('删除重复行：只点了一个单元格时自动选中整块数据',async()=>{
  const region={Rows:{Count:3},Columns:{Count:1},MergeCells:false,HasFormula:false,Value2:[['a'],['b'],['a']],Cells:{Item:()=>({Resize:()=>({set Value2(v){region.written=v;},ClearContents(){}})})}};
  let selected=null;region.Select=()=>{selected=region;};
  const cell={Rows:{Count:1},Columns:{Count:1},CurrentRegion:region};
  const a={ActiveWorkbook:{Name:'s.xlsx'},ActiveSheet:{Name:'S'},get Selection(){return selected||cell;}};
  const msg=await bridge(a,'et').T.run('removeDuplicates',{host:'et'});
  assert.equal(selected,region);assert.equal(JSON.stringify(region.written),JSON.stringify([['a'],['b']]));assert.match(msg,/删除 1 行/);
});

// WPS 文字的金额转大写：用字符位置模拟文档。
function wordDoc(text,sel){
  const doc={text};const range=(s,e)=>({Start:s,End:e,get Text(){return doc.text.slice(s,e);},set Text(v){doc.text=doc.text.slice(0,s)+v+doc.text.slice(e);},InsertAfter(v){doc.text=doc.text.slice(0,e)+v+doc.text.slice(e);}});
  const d={Range:(s,e)=>range(s,e),Name:'a.docx'};
  const s={Start:sel[0],End:sel[1],get Text(){return doc.text.slice(sel[0],sel[1]);},get Range(){return range(sel[0],sel[1]);}};
  return {a:{ActiveDocument:d,Selection:s},doc};
}
test('金额大写：WPS 双击只选到小数点前，也按完整金额 128560.50 计算，并填进「大写：」空位',async()=>{
  const t='本合同总金额为人民币 128560.50 元（大写：　　　　）。';const at=t.indexOf('128560');
  const {a,doc}=wordDoc(t,[at,at+6]);const msg=await bridge(a,'wps').T.run('rmbUpper',{host:'wps'});
  assert.equal(doc.text,'本合同总金额为人民币 128560.50 元（大写：壹拾贰万捌仟伍佰陆拾元伍角）。');assert.match(msg,/已填入大写/);
});
test('金额大写：没有空位时加在完整数字后面；已经有大写时不重复加；光标放在数字里也行',async()=>{
  const t='首付 38568.15 元，余款另付。',at=t.indexOf('38568');
  const one=wordDoc(t,[at+2,at+2]);await bridge(one.a,'wps').T.run('rmbUpper',{host:'wps'});
  assert.equal(one.doc.text,'首付 38568.15（叁万捌仟伍佰陆拾捌元壹角伍分） 元，余款另付。');
  const two=wordDoc('合计 100（壹佰元整）元',[3,6]);const msg=await bridge(two.a,'wps').T.run('rmbUpper',{host:'wps'});
  assert.match(msg,/已经有大写/);assert.equal(two.doc.text,'合计 100（壹佰元整）元');
});

// 正文排版：段落模拟。
function paragraphs(list,{start=0,end=1}={}){
  const made=list.map(p=>{const r={Text:p.text+'\r',Font:{Size:p.size||12,Bold:p.bold==='mixed'?9999999:!!p.bold},Characters:{Item:()=>({Font:{Bold:p.bold==='mixed'?-1:!!p.bold}})},ParagraphFormat:{},Style:p.style||'正文',deleted:false,Information:()=>false,Delete(){r.deleted=true;}};return {OutlineLevel:p.level||10,Range:r};});
  const s={Start:start,End:end,Text:list.map(p=>p.text).join('\r'),Paragraphs:{Count:made.length,Item:i=>made[i-1]},Range:{},Font:{},ParagraphFormat:{}};
  return {a:{ActiveDocument:{Name:'a.docx'},Selection:s},made};
}
test('正文排版：没选中文字时不改文档，提示先选中',async()=>{
  const {a}=paragraphs([{text:'部门周例会纪要',size:18}],{start:5,end:5});
  await assert.rejects(bridge(a,'wps').T.run('bodyFormat',{host:'wps'}),/请先选中要整理的正文/);
});
test('正文排版：全选后标题跳过、空行删除、只确认一次并说明改了几段',async()=>{
  const {a,made}=paragraphs([{text:'部门周例会纪要',size:18},{text:''},{text:'时间：11 月 20 日'},{text:'一、上周工作回顾',bold:true},{text:'1. 完成了报价单'},{text:'   '},{text:'二级标题',level:2}]);
  const {T,asked}=bridge(a,'wps');const msg=await T.runFromUi('bodyFormat',{host:'wps'});
  assert.equal(asked.length,1);assert.match(asked[0],/整理 2 段正文.*跳过 3 个标题.*删除 2 个空行/);
  assert.match(msg,/已整理 2 段正文.*跳过 3 个标题.*删除 2 个空行/);
  assert.equal(made[0].Range.Font.Size,18,'标题字号不变');assert.equal(made[2].Range.ParagraphFormat.SpaceBefore,0,'段前距清零（真机：原文段前空得忽大忽小）');assert.equal(made[2].Range.Font.Size,12);assert.equal(made[2].Range.Font.NameFarEast,'宋体');assert.equal(made[2].Range.Font.NameAscii,'宋体','数字和英文也换成同一字体（真机：Arial / Times 的数字把行撑高）');
  assert.ok(made[1].Range.deleted&&made[5].Range.deleted);assert.ok(!made[2].Range.deleted);
});
test('正文排版（真机）：整段加粗返回“混合”时按首字判断；没加粗的短“二、本周计划”也算小标题，长的“一、……。”句子仍是正文',async()=>{
  const {a,made}=paragraphs([{text:'一、上周工作回顾',bold:'mixed'},{text:'二、本周计划'},{text:'三、各部门请于本周五前提交材料，逾期不再受理。'},{text:'1、整理年终总结材料'}]);
  const {T,asked}=bridge(a,'wps');await T.runFromUi('bodyFormat',{host:'wps'});
  assert.match(asked[0],/整理 2 段正文.*跳过 2 个标题/);
  assert.equal(made[0].Range.Font.NameFarEast,undefined);assert.equal(made[1].Range.Font.NameFarEast,undefined);assert.equal(made[2].Range.Font.NameFarEast,'宋体');
});

test('打开生成的文件：从 WPS 发起时按扩展名用 WPS 程序打开，不依赖文件关联；找不到 WPS 再按系统默认',async t=>{
  assert.equal(wpsProgramFor('C:/a/报价单.xlsx'),'et.exe');assert.equal(wpsProgramFor('C:/a/x.docx'),'wps.exe');assert.equal(wpsProgramFor('C:/a/x.pptx'),'wpp.exe');assert.equal(wpsProgramFor('C:/a/x.pdf'),null);
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'wps-'));t.after(()=>fs.rmSync(dir,{recursive:true,force:true}));fs.mkdirSync(path.join(dir,'office6'));fs.writeFileSync(path.join(dir,'office6','et.exe'),'');
  const calls=[];const start=async(cmd,args)=>{calls.push([path.basename(cmd),args[0]]);};
  const file=path.join(dir,'报价单.xlsx');
  await openDocument(file,{prefer:'wps',platform:'win32',findRoot:async()=>dir,start});
  await openDocument(file,{prefer:'wps',platform:'win32',findRoot:async()=>null,start});
  await openDocument(file,{platform:'win32',findRoot:async()=>dir,start});
  assert.deepEqual(calls,[['et.exe',file],['explorer.exe',file],['explorer.exe',file]]);
});

test('生成文件的文件夹：和备份分开；中文、日文默认名不同；可自定义，必须是完整路径',()=>{
  const s=defaultSettings('/b');assert.equal(s.outputs.directory,'');
  assert.equal(outputDirectory({...s,ui:{language:'zh'}},{home:'/h'}),path.join('/h','Documents','办公工具箱文件'));
  assert.equal(outputDirectory({...s,ui:{language:'ja'}},{home:'/h'}),path.join('/h','Documents','オフィスツールボックス'));
  assert.equal(outputDirectory({...s,ui:{language:'auto'}},{home:'/h',locale:'ja-JP'}),path.join('/h','Documents','オフィスツールボックス'));
  assert.equal(outputDirectory({...s,outputs:{directory:'/x/成品'}},{home:'/h'}),'/x/成品');
  assert.throws(()=>validateSettings({outputs:{directory:'相对路径'}},s),/生成文件夹/);assert.equal(validateSettings({outputs:{directory:''}},s).outputs.directory,'');
});

test('商务模板：示例行不再带数量和单价，合计从 0 开始，避免忘删多报金额',async()=>{
  for(const f of ['quote-zh','quote-ja','invoice-ja','receipt-zh']){const wb=new ExcelJS.Workbook();await wb.xlsx.readFile(new URL('../public/templates/business/'+f+'.xlsx',import.meta.url).pathname.replace(/^\/([A-Za-z]:)/,'$1'));const ws=wb.worksheets[0];
    assert.equal(ws.getCell('D11').value,null,f);assert.equal(ws.getCell('E11').value,null,f);assert.match(String(ws.getCell('B11').value),/［/,f);}
});

test('新手模式：设置项默认关闭，只接受开/关；引导最后一页让客户选择',()=>{
  const s=defaultSettings('/b');assert.equal(s.ui.beginner,false);assert.equal(validateSettings({ui:{beginner:true}},s).ui.beginner,true);assert.throws(()=>validateSettings({ui:{beginner:'yes'}},s),/新手模式/);
  const src=fs.readFileSync(new URL('../public/welcome.js',import.meta.url),'utf8');assert.match(src,/只显示常用工具（推荐新手）/);assert.match(src,/update\(\{ui:\{beginner:!!on\}\}\)/);
  const qp=fs.readFileSync(new URL('../public/quick-panel.js',import.meta.url),'utf8');assert.match(qp,/var BEGINNER=\{wps:\[/);assert.match(qp,/显示全部工具/);
});
test('PPT 版式工作室窗口按屏幕大小打开，高分屏不再是一个小窗',()=>{
  const ctx={window:{}};vm.runInNewContext(fs.readFileSync(new URL('../public/studio-launch.js',import.meta.url),'utf8'),ctx);const size=ctx.window.StudioLaunch.dialogSize;
  assert.deepEqual([...size({availWidth:1707,availHeight:1067},1.5)],[2176,1408]);assert.deepEqual([...size({availWidth:1024,availHeight:700},1)],[1040,760]);
});
test('没接 AI 时：按问题推荐本地工具，并给“去设置 AI”按钮',()=>{
  const src=fs.readFileSync(new URL('../public/ai-chat-panel.js',import.meta.url),'utf8');assert.match(src,/suggestLocal\(q\)/);assert.match(src,/QuickPanel\.suggest\(q,3\)/);assert.match(src,/去设置 AI/);
  const qp=fs.readFileSync(new URL('../public/quick-panel.js',import.meta.url),'utf8');assert.match(qp,/window\.QuickPanel=\{suggest:/);assert.match(qp,/removeDuplicates:'重复 去重/);
});
test('导出 PDF 用原文件名，重名才加 -2；工作室 PPT 用标题命名、放在生成文件夹',()=>{
  const t=fs.readFileSync(new URL('../public/tools/transfer.js',import.meta.url),'utf8');assert.match(t,/freePdfName/);assert.doesNotMatch(t,/name\+'-'\+Date\.now\(\)\+'\.pdf'/);
  const s=fs.readFileSync(new URL('../server.mjs',import.meta.url),'utf8');assert.match(s,/saveNew\(studioDir\|\|await outputDir\(\)/);
});
test('AI 未接入时的工具推荐（真机）：QuickPanel 只导出一次，suggest 不会被后面覆盖；日文“～たい”也算口头语',async()=>{
  const src=fs.readFileSync(new URL('../public/quick-panel.js',import.meta.url),'utf8');
  assert.equal(src.match(/window\.QuickPanel=/g).length,1);assert.match(src,/window\.QuickPanel=\{suggest:function/);
  assert.match(src,/したい\|たい\|ほしい/);
});
test('i18n：脚本写进页面的日文提示，切回中文界面时能还原成中文',()=>{
  const root=path.resolve(new URL('..',import.meta.url).pathname.replace(/^\/([A-Za-z]:)/,'$1'));
  const ctx={};ctx.window=ctx;ctx.localStorage={getItem:()=>'ja',setItem(){}};ctx.navigator={language:'ja'};ctx.setTimeout=()=>0;ctx.setInterval=()=>0;
  const node={nodeType:3,nodeValue:'',parentNode:null};
  const body={nodeType:1,tagName:'BODY',getAttribute:()=>null,hasAttribute:()=>false,ownerDocument:null};node.parentNode=body;
  const doc={documentElement:{},body,title:'',addEventListener(){},querySelectorAll:()=>[],createTreeWalker:()=>{let done=false;return {nextNode:()=>done?null:(done=true,node)};}};body.ownerDocument=doc;ctx.document=doc;
  for(const f of ['i18n-ja.js','i18n-ja-extra.js','i18n.js'])vm.runInNewContext(fs.readFileSync(path.join(root,'public',f),'utf8'),ctx);
  const I=ctx.OfficeI18n;I.use('ja');node.nodeValue=I.t('尚未设置版本信息地址。');assert.match(node.nodeValue,/[ぁ-ん]/);
  I.use('zh');assert.equal(node.nodeValue,'尚未设置版本信息地址。');
  I.use('ja');assert.equal(node.nodeValue,I.t('尚未设置版本信息地址。'));
});
test('正文排版：字体按文档内容选，日文正文用 Yu Mincho，中文正文用宋体（不看界面语言）',async()=>{
  const ja=paragraphs([{text:'本日の会議では、来期の予算について話し合いました。'},{text:'次回は来週の月曜日です。'}]);
  await bridge(ja.a,'wps').T.runFromUi('bodyFormat',{host:'wps'});assert.equal(ja.made[0].Range.Font.NameFarEast,'Yu Mincho');
  const zh=paragraphs([{text:'本周会议讨论了下季度预算。'},{text:'下次会议在周一。'}]);
  await bridge(zh.a,'wps').T.runFromUi('bodyFormat',{host:'wps'});assert.equal(zh.made[0].Range.Font.NameFarEast,'宋体');
});
test('功能区：「打开工具箱」放最前，和「PPT 版式工作室」隔开；成功提示带 ✓',()=>{
  const xml=fs.readFileSync(new URL('../public/ribbon.xml',import.meta.url),'utf8');
  const g=xml.slice(xml.indexOf('<group id="toolsGroup"'),xml.indexOf('</group>',xml.indexOf('<group id="toolsGroup"')));
  assert.match(g,/^<group id="toolsGroup"[^>]*><button id="toolsOpen"/);
  assert.ok(g.indexOf('id="studioOpen"')>g.lastIndexOf('<separator'),'工作室在最后一条分隔线之后');
  assert.doesNotMatch(g.match(/<button[^>]*id="studioOpen"[^>]*>/)[0],/size="large"/);
  const main=fs.readFileSync(new URL('../public/main.js',import.meta.url),'utf8'),i18n=fs.readFileSync(new URL('../public/i18n.js',import.meta.url),'utf8');
  assert.match(main,/function done\(message\)/);assert.match(i18n,/OfficeI18n\.done=function\(message\)\{[^}]*'✓ '\+translate\(message\)/);
});
