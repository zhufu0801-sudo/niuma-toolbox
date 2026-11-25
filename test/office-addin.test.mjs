import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import https from 'node:https';
import {classifyOffice} from '../lib/windows.mjs';
import {createOfficeServer,officeFileAllowed} from '../lib/office-server.mjs';
const read=f=>fs.readFileSync(new URL('../'+f,import.meta.url),'utf8');
function tools(){const w={};vm.runInNewContext(read('public/office/office-tools.js'),{window:w,Promise});return w.OfficeAddinTools;}
// 模拟 Excel：选区是二维数组，getCell 返回可写对象；load/sync 不做事
function excel(values,formulas){const writes=[];const fmt={font:{},fill:{},load(){},autofitColumns(){fmt.autofit=true;},autofitRows(){}};
 const r={rowCount:values.length,columnCount:values[0].length,values,formulas:formulas||values.map(row=>row.map(v=>v??'')),format:fmt,load(){},getRow(){return {format:{font:{},fill:{}}};},getCell(i,j){return {set values(v){writes.push([i,j,v[0][0]]);values[i][j]=v[0][0];},load(){},left:j*50,top:i*20,width:50,height:20};}};
 const ctx={workbook:{getSelectedRange:()=>r,worksheets:{getActiveWorksheet:()=>ctx.sheet}},sheet:{shapes:{items:[],load(){}}},sync:async()=>{}};
 return {r,writes,ctx,api:{Excel:{run:fn=>fn(ctx)}}};}
function word(pics,page){const list={items:pics.map(([w,h])=>({width:w,height:h})),load(){}};
 const ctx={document:{getSelection:()=>({inlinePictures:list,paragraphs:{items:[{},{}],load(){}}}),sections:{getFirst:()=>{if(!page)throw Error('ApiNotFound');return {pageSetup:{...page,load(){}}};}}},sync:async()=>{}};
 return {list,ctx,api:{Word:{run:fn=>fn(ctx)}}};}
test('Office 检测：区分 Microsoft 365、买断版、MSI 版和未安装',()=>{
 assert.equal(classifyOffice({version:'16.0.1',productIds:'O365ProPlusRetail'}).name,'Microsoft 365');
 assert.equal(classifyOffice({version:'16.0.1',productIds:'ProPlus2021Volume'}).name,'Office 2021');
 const old=classifyOffice({msi:[{ver:'15.0',app:'Excel'}]});assert.equal(old.supported,false);assert.equal(old.name,'Office 2013');
 assert.equal(classifyOffice({}).installed,false);
});
test('Office 版 Excel 工具：跨列居中检查、空白填上一格先确认再执行、清理空格跳过公式',async()=>{
 const T=tools();
 const a=excel([['标题',null,null]]);assert.match(await T.run('excel','centerAcross',a.api),/跨列居中/);assert.equal(a.r.format.horizontalAlignment,'CenterAcrossSelection');
 const b=excel([['标题','备注']]);await assert.rejects(T.run('excel','centerAcross',b.api),/多个非空/);
 const c=excel([['销售'],[null],['财务'],[null]]);const ask=await T.run('excel','fillBlanks',c.api);assert.match(ask.confirm,/2 个空白/);assert.equal(c.writes.length,0);
 assert.match(await T.run('excel','fillBlanks',c.api,{confirmed:true}),/已填充 2/);assert.deepEqual(c.r.values.map(x=>x[0]),['销售','销售','财务','财务']);
 const d=excel([[' a '],['=A1 ']],[[' a '],['=A1 ']]);assert.match(await T.run('excel','trimCells',d.api),/1 个/);assert.deepEqual(d.writes,[[0,0,'a']]);
 const e=excel([['x']]);e.r.format.wrapText=true;assert.match(await T.run('excel','shrinkToFit',e.api),/换行已关闭/);assert.equal(e.r.format.shrinkToFit,true);
});
test('Office 版 Excel 图片适应单元格：只处理左上角在选区内的图片',async()=>{
 const T=tools();const x=excel([[1,2],[3,4]]);const pic={type:'Image',left:52,top:2,width:100,height:100},far={type:'Image',left:500,top:500,width:10,height:10};x.ctx.sheet.shapes.items=[pic,far];
 assert.match(await T.run('excel','fitPicturesToCells',x.api),/1 张/);assert.equal(pic.width,16);assert.equal(pic.height,16);assert.equal(pic.left,67);assert.equal(pic.placement,'TwoCell');assert.equal(far.width,10);
});
test('Office 版 Word 图片工具：统一大小、超出版心、证件照',async()=>{
 const T=tools();const a=word([[200,100],[400,400]]);assert.match(await T.run('word','unifyPictures',a.api),/2 张/);assert.deepEqual(a.list.items.map(p=>[p.width,p.height]),[[200,100],[200,200]]);
 const b=word([[600,300]]);assert.match(await T.run('word','fitPictures',b.api),/按 A4/);assert.equal(b.list.items[0].width,425);
 const c=word([[600,300]],{pageWidth:595,leftMargin:72,rightMargin:72});assert.doesNotMatch(await T.run('word','fitPictures',c.api),/按 A4/);assert.equal(c.list.items[0].width,451);
 const d=word([[317,200]]);await T.run('word','idCardSize',d.api);assert.equal(Math.round(d.list.items[0].width*100)/100,242.65);
 await assert.rejects(T.run('word','idCardSize',word([[300,300]]).api),/长宽比/);
});
test('Office 清单模板：结构完整，所有地址都走本机 https，图标存在',()=>{
 const xml=read('public/office/manifest-template.xml');
 for(const k of ['<Id>{{ID}}</Id>','<Version>{{VERSION}}</Version>','<Host Name="Workbook"/>','<Host Name="Document"/>','<Host Name="Presentation"/>','xsi:type="ShowTaskpane"','Locale="ja-JP"'])assert.ok(xml.includes(k),k);
 const urls=[...xml.matchAll(/DefaultValue="([^"]*(?:\/office\/)[^"]*)"/g)].map(m=>m[1]);assert.ok(urls.length>=6);for(const u of urls)assert.match(u,/^\{\{BASE\}\}\/office\/[a-z0-9-]+\.(html|png)$/);
 for(const s of [16,32,80])assert.ok(fs.existsSync(new URL('../public/office/icon-'+s+'.png',import.meta.url)));
 assert.match(read('public/office/taskpane.html'),/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js/);
});
test('Office https 服务：只认本机地址，只给侧栏文件',async t=>{
 assert.ok(officeFileAllowed('/office/taskpane.html'));assert.ok(officeFileAllowed('/office/shortcuts.json'));assert.ok(!officeFileAllowed('/office/secrets.json'));assert.ok(officeFileAllowed('/i18n.js'));assert.ok(!officeFileAllowed('/api/settings'));assert.ok(!officeFileAllowed('/office/../server.mjs'));
 const pfx=fs.readFileSync(new URL('./fixtures/test-only-localhost.pfx',import.meta.url));
 const server=createOfficeServer({pfx,passphrase:'test-only',publicDir:new URL('../public',import.meta.url).pathname,port:0});await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>server.close());
 const port=server.address().port;const get=(path,host)=>new Promise((resolve,reject)=>{https.get({host:'127.0.0.1',port,path,rejectUnauthorized:false,headers:{Host:host}},r=>{r.resume();resolve(r.statusCode);}).on('error',reject);});
 // 测试用 port:0，服务按 0 比对 Host；这里只确认拒绝外部主机名
 assert.equal(await get('/office/taskpane.html','evil.example'),403);
});
test('安装与卸载脚本：识别两套软件、让用户选择、Office 登记和证书都能清理',()=>{
 const init=read('scripts/package/Initialize.ps1'),un=read('scripts/package/Uninstall.ps1');
 for(const k of ['Find-Wps','Find-Office','Choose-Targets','ClickToRun\\Configuration','WEF\\Developer','Cert:\\CurrentUser\\Root','manifest-template.xml','targets.json',"Say '"])assert.ok(init.includes(k),k);
 for(const k of ['WEF\\Developer','office-cert.thumb','Cert:\\CurrentUser\\Root','WpsCoachSheets','ai.json'])assert.ok(un.includes(k),k);
 assert.ok(init.charCodeAt(0)===0xFEFF&&un.charCodeAt(0)===0xFEFF,'PowerShell 5 需要 BOM 才能正确读中文/日文');
});
test('Office 版补齐的工具：查找替换、清理空段落、文字转表格、按大纲生成 PPT、筛选开关',async()=>{
 const T=tools();
 const x=excel([['a-b'],['=A1'],[5]],[['a-b'],['=A1'],[5]]);assert.match(await T.run('excel','replaceText',x.api,{find:'-',replace:'+'}),/1 个/);assert.deepEqual(x.writes,[[0,0,'a+b']]);
 await assert.rejects(T.run('excel','replaceText',x.api,{find:''}),/查找/);
 const deleted=[];const ps={items:[{text:'正文',delete(){deleted.push(0);}},{text:'  ',delete(){deleted.push(1);}},{text:'',delete(){deleted.push(2);}}],load(){}};
 const hits={items:[{insertText(t,m){hits.done=[t,m];}}],load(){}};let table=null;
 const wctx={document:{getSelection:()=>({paragraphs:ps,text:'姓名\t部门\n甲\t销售\n',load(){},search:()=>hits,insertTable(r,c,where,v){table={r,c,where,v};}})},sync:async()=>{}};const wapi={Word:{run:fn=>fn(wctx)}};
 assert.match(await T.run('word','cleanEmptyParagraphs',wapi),/2 个/);assert.deepEqual(deleted,[1,2]);
 assert.match(await T.run('word','replaceText',wapi,{find:'甲',replace:'乙'}),/1 处/);assert.deepEqual(hits.done,['乙','Replace']);
 assert.match(await T.run('word','textToTable',wapi),/2 行 × 2 列/);assert.equal(table.where,'After');assert.equal(JSON.stringify(table.v[1]),JSON.stringify(['甲','销售']));
 const slides=[];const pctx={presentation:{slides:{items:slides,load(){},add(){const shapes=[];slides.push({shapes:{addTextBox(t){const tb={text:t,textFrame:{textRange:{font:{}}}};shapes.push(tb);return tb;},list:shapes}});}}},sync:async()=>{}};
 assert.match(await T.run('powerpoint','makeSlides',{PowerPoint:{run:fn=>fn(pctx)}},{text:'概览\n目标\n计划\n\n下一步\n分工'}),/2 页/);assert.equal(slides.length,2);assert.equal(slides[1].shapes.list[0].text,'下一步');assert.equal(slides[0].shapes.list[1].text,'• 目标\n• 计划');
 const f=excel([['h'],[1]]);let applied=false;f.ctx.workbook.worksheets.getActiveWorksheet=()=>({autoFilter:{enabled:false,load(){},apply(){applied=true;},remove(){}}});assert.match(await T.run('excel','filterCells',f.api),/打开/);assert.ok(applied);
});
test('Office 版 PowerPoint：统一本页字体保留标题大小，复制当前页插到原页之后',async()=>{
 const T=tools();
 const box=(name,type,size,font,text=true)=>{const f={size,name:font,load(){}};return {name,type,textFrame:{hasText:text,load(){},textRange:{font:f}},f};};
 const shapes=[box('标题 1','Placeholder',40,'黑体'),box('文本框 2','TextBox',18,'微软雅黑'),box('文本框 3','TextBox',24,'微软雅黑'),box('文本框 4','TextBox',18,'宋体'),{name:'图片 5',type:'Image',get textFrame(){throw Error('图片不该读文字');}}];
 let inserted=null;const slide={id:'257#',shapes:{items:shapes,load(){}},exportAsBase64:()=>({value:'QUJD'})};
 const ctx={presentation:{getSelectedSlides:()=>({items:[slide],load(){}}),insertSlidesFromBase64:(b,o)=>{inserted=[b,o];}},sync:async()=>{}};
 const api={PowerPoint:{run:fn=>fn(ctx)}};
 assert.match(await T.run('powerpoint','formatSlide',api),/4 个文字框为“微软雅黑”，正文统一为 18 号，标题保持原有大小/);
 assert.equal(shapes[0].f.size,40);assert.deepEqual(shapes.slice(1,4).map(s=>[s.f.size,s.f.name]),[[18,'微软雅黑'],[18,'微软雅黑'],[18,'微软雅黑']]);
 assert.match(await T.run('powerpoint','duplicateSlide',api),/副本/);assert.equal(JSON.stringify(inserted),JSON.stringify(['QUJD',{targetSlideId:'257#',formatting:'KeepSourceFormatting'}]));
 const none={presentation:{getSelectedSlides:()=>({items:[],load(){}})},sync:async()=>{}};await assert.rejects(T.run('powerpoint','duplicateSlide',{PowerPoint:{run:fn=>fn(none)}}),/点一下/);
 assert.equal(T.list('powerpoint').map(x=>x.id).join(),'makeSlides,formatSlide,unifyDeckFont,removeAnimations,unifyTitles,arrangeVisuals,slideNumbers,bodySpacing,applyDesign,removeNotes,duplicateSlide,backupFile,exportPdf');
});
test('Office 版 Word：目录先更新已有，没有标题时拒绝；匹配周围格式跳过标题段',async()=>{
 const T=tools();let updated=0,inserted=null;
 const mk=(fields,styles)=>{const sel={paragraphs:{items:[],load(){}},insertField:(loc,type,text)=>{inserted=[loc,type,text];return {updateResult(){updated++;}};}};
  return {Word:{run:fn=>fn({document:{body:{fields:{items:fields,load(){}},paragraphs:{items:styles.map(s=>({styleBuiltIn:s})),load(){}}},getSelection:()=>sel},sync:async()=>{}})}};};
 assert.equal(await T.run('word','toc',mk([{type:'Toc',updateResult(){updated++;}}],[])),'目录已更新。');assert.equal(updated,1);
 await assert.rejects(T.run('word','toc',mk([],['Normal'])),/一至三级标题/);
 assert.match(await T.run('word','toc',mk([],['Normal','Heading2'])),/插入目录/);assert.equal(inserted.join('|'),'End|TOC|\\o "1-3" \\h \\z \\u');
 const para=(o)=>({isNullObject:false,text:'x',load(){},font:{name:'宋体',size:12,color:'#000000',load(){}},...o});
 const prev=para({styleBuiltIn:'Heading1',style:'标题 1',font:{name:'黑体',size:22,load(){}}}),next=para({styleBuiltIn:'Normal',style:'正文',lineSpacing:20,spaceAfter:6,spaceBefore:0,firstLineIndent:24});
 const mine={load(){},getPreviousOrNullObject:()=>prev,getNextOrNullObject:()=>next};const sel={paragraphs:{items:[mine],load(){}},font:{},load(){}};
 const msg=await T.run('word','matchFormat',{Word:{run:fn=>fn({document:{getSelection:()=>sel},sync:async()=>{}})}});
 assert.match(msg,/宋体.*12 号/);assert.equal(mine.style,'正文');assert.equal(mine.firstLineIndent,24);assert.equal(sel.font.name,'宋体');
});
test('Office https 服务：内容卡片只转发读取和带口令的保存，拒绝别的来源',async t=>{
 const http=await import('node:http'),net=await import('node:net');const seen=[];
 const api=http.createServer((req,res)=>{let b='';req.on('data',c=>b+=c);req.on('end',()=>{seen.push({m:req.method,p:req.url,host:req.headers.host,tok:req.headers['x-coach-token'],origin:req.headers.origin,b});res.writeHead(200,{'Content-Type':'application/json'});res.end(req.url==='/api/session'?'{"token":"abc"}':'{"cards":[]}');});});
 await new Promise(r=>api.listen(0,'127.0.0.1',r));t.after(()=>api.close());
 const free=await new Promise(r=>{const s=net.createServer().listen(0,'127.0.0.1',()=>{const p=s.address().port;s.close(()=>r(p));});});
 const pfx=fs.readFileSync(new URL('./fixtures/test-only-localhost.pfx',import.meta.url));
 const server=createOfficeServer({pfx,passphrase:'test-only',publicDir:new URL('../public',import.meta.url).pathname,port:free,apiPort:api.address().port});await new Promise(r=>server.listen(free,'127.0.0.1',r));t.after(()=>server.close());
 const call=(method,path,headers={},body)=>new Promise((resolve,reject)=>{const q=https.request({host:'127.0.0.1',port:free,method,path,rejectUnauthorized:false,headers:{Host:'localhost:'+free,...headers}},r=>{let d='';r.on('data',c=>d+=c);r.on('end',()=>resolve({code:r.statusCode,d,cors:r.headers['access-control-allow-origin']}));});q.on('error',reject);q.end(body);});
 const s=await call('GET','/api/session');assert.equal(s.code,200);assert.equal(JSON.parse(s.d).token,'abc');assert.equal(s.cors,undefined);
 assert.equal((await call('GET','/api/cards',{'X-Coach-Token':'abc'})).code,200);
 assert.equal((await call('POST','/api/cards',{'Content-Type':'application/json','X-Coach-Token':'abc'},'{"block":{"type":"text","text":"x"},"sourceHost":"wps"}')).code,200);
 assert.equal((await call('POST','/api/cards',{'Content-Type':'application/json'},'{}')).code,403);
 assert.equal((await call('DELETE','/api/cards',{'X-Coach-Token':'abc'})).code,405);
 assert.equal((await call('GET','/api/cards',{Origin:'https://evil.example','X-Coach-Token':'abc'})).code,403);
 assert.equal((await call('GET','/api/settings')).code,404);
 assert.equal(seen.length,3);assert.ok(seen.every(x=>x.host==='127.0.0.1:'+api.address().port&&!x.origin));assert.equal(seen[2].b,'{"block":{"type":"text","text":"x"},"sourceHost":"wps"}');
});
test('Office 版内容卡片：Excel 存表格、插入前提示覆盖、公式样文字保持为文字；Word 插入表格',async()=>{
 const T=tools();const posts=[];const fetcher=async(url,o={})=>{if(url==='/api/session')return {ok:true,json:async()=>({token:'k'})};if(o.method==='POST'){posts.push(JSON.parse(o.body));return {ok:true,json:async()=>({card:{title:'2 行 × 2 列表格：姓名、金额'}})};}return {ok:true,json:async()=>({cards:[{id:'1'}]})};};
 const a=excel([['姓名','金额'],['张三',5]]);assert.match(await T.run('excel','saveCard',a.api,{fetch:fetcher}),/已设为内容模块：2 行/);assert.equal(JSON.stringify(posts[0]),JSON.stringify({block:{type:'table',rows:[['姓名','金额'],['张三','5']]},sourceHost:'et'}));
 assert.equal((await T.cards(fetcher)).length,1);
 const target={values:[['',''],['旧','']],load(){}};const sel={getCell:()=>({getResizedRange:(h,w)=>{target.size=[h,w];return target;}})};const ctx={workbook:{getSelectedRange:()=>sel},sync:async()=>{}};const api={Excel:{run:fn=>fn(ctx)}};
 const card={title:'测试',block:{type:'table',rows:[['=1+1','b'],['c','d']]}};
 const ask=await T.run('excel','insertCard',api,{card});assert.match(ask.confirm,/1 个单元格/);
 assert.match(await T.run('excel','insertCard',api,{card,confirmed:true}),/已插入内容卡片：测试/);assert.equal(JSON.stringify(target.values),JSON.stringify([["'=1+1",'b'],['c','d']]));assert.equal(target.size.join(),'1,1');
 let tbl=null;const wctx={document:{getSelection:()=>({insertTable:(r,c,loc,v)=>{tbl=[r,c,loc,v];},insertText(){}})},sync:async()=>{}};
 await T.run('word','insertCard',{Word:{run:fn=>fn(wctx)}},{card});assert.equal(JSON.stringify(tbl),JSON.stringify([2,2,'After',[['=1+1','b'],['c','d']]]));
 await assert.rejects(T.run('word','saveCard',{Word:{run:fn=>fn({document:{getSelection:()=>({tables:{items:[],load(){}},text:'  ',load(){}})},sync:async()=>{}})}},{fetch:fetcher}),/先选中/);
});
test('Office 版备份与导出 PDF：只备份本机文件；Word 分片生成 PDF 交给本机服务；Excel 给出菜单指引',async()=>{
 const T=tools();const posts=[];const fetcher=async(url,o={})=>{if(url==='/api/session')return {ok:true,json:async()=>({token:'k'})};posts.push([url,JSON.parse(o.body),o.headers['X-Coach-Token']]);return {ok:true,json:async()=>(url==='/api/backup'?{backupPath:'D:\\备份\\a-备份.docx'}:{path:'D:\\备份\\a-20261111.pdf'})};};
 const office=(url,parts)=>({context:{document:{url,getFileAsync:(type,opt,cb)=>{assert.equal(type,'pdf');assert.equal(opt.sliceSize,4194304);let closed=0;cb({status:'succeeded',value:{sliceCount:parts.length,getSliceAsync:(i,c)=>c({status:'succeeded',value:{data:parts[i]}}),closeAsync:()=>closed++}});}}},FileType:{Pdf:'pdf'}});
 assert.match(await T.run('word','backupFile',{Office:office('C:\\文档\\a.docx',[])},{fetch:fetcher}),/已备份到：D:\\备份\\a-备份\.docx\n备份的是最近一次保存的内容/);assert.equal(JSON.stringify(posts[0]),JSON.stringify(['/api/backup',{path:'C:\\文档\\a.docx'},'k']));
 await assert.rejects(T.run('excel','backupFile',{Office:office('https://d.docs.live.net/x/a.xlsx',[])},{fetch:fetcher}),/OneDrive/);
 await assert.rejects(T.run('powerpoint','backupFile',{Office:office('',[])},{fetch:fetcher}),/先把文件保存/);
 const pdf=[...Buffer.from('%PDF-1.7 ')].concat([200,201]);
 assert.match(await T.run('word','exportPdf',{Office:office('C:\\文档\\方案.docx',[pdf.slice(0,5),pdf.slice(5)])},{fetch:fetcher}),/PDF 已导出：D:\\备份/);
 const sent=posts.at(-1);assert.equal(sent[0],'/api/save-pdf');assert.equal(sent[1].name,'方案.docx');assert.deepEqual([...Buffer.from(sent[1].data,'base64')],pdf);
 await assert.rejects(T.run('excel','exportPdf',{Office:office('C:\\a.xlsx',[])},{fetch:fetcher}),/创建 PDF\/XPS/);
});
test('Office 保存与另存后读取最新位置；支持 Mac 本机路径和 file URL，不回退到陈旧地址',async()=>{
 const T=tools(),doc={url:'C:\\old.docx',getFilePropertiesAsync:cb=>cb({status:'succeeded',value:{url:'file:///C:/new%20name.docx'}})},api={Office:{context:{document:doc}}};
 assert.equal(T.localPath(await T.documentUrl(api)),'C:/new name.docx');
 doc.getFilePropertiesAsync=cb=>cb({status:'succeeded',value:{url:'/Users/test/日本語.docx'}});assert.equal(T.localPath(await T.documentUrl(api)),'/Users/test/日本語.docx');
 assert.equal(T.localPath('file://localhost/Users/test/a%20b.docx'),'/Users/test/a b.docx');assert.equal(T.localPath('https://example.test/a.docx'),'');assert.equal(T.localPath('relative.docx'),'');
 doc.getFilePropertiesAsync=cb=>cb({status:'succeeded',value:{url:''}});assert.equal(await T.documentUrl(api),'');
 doc.getFilePropertiesAsync=cb=>cb({status:'failed'});await assert.rejects(T.documentUrl(api),/无法读取当前文件/);
});
test('本机服务 /api/save-pdf：要口令、只收 PDF、存到备份文件夹且不覆盖',async t=>{
 const os=await import('node:os'),path=await import('node:path');const {createCoachServer}=await import('../server.mjs');
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'pdf-'));t.after(()=>fs.rmSync(dir,{recursive:true,force:true}));
 const server=createCoachServer({backupDir:path.join(dir,'备份'),cardsFile:path.join(dir,'cards.json'),settingsFile:path.join(dir,'settings.json')});await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>server.close());
 const base='http://127.0.0.1:'+server.address().port,tok=(await (await fetch(base+'/api/session')).json()).token;
 const send=(body,k=tok)=>fetch(base+'/api/save-pdf',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':k},body:JSON.stringify(body)});
 assert.equal((await send({name:'a',data:Buffer.from('%PDF-x').toString('base64')},'bad')).status,403);
 assert.equal((await send({name:'a',data:Buffer.from('hello').toString('base64')})).status,422);
 const one=await (await send({name:'..\\..\\方案:v1.docx',data:Buffer.from('%PDF-1').toString('base64')})).json(),two=await (await send({name:'..\\..\\方案:v1.docx',data:Buffer.from('%PDF-2').toString('base64')})).json();
 assert.ok(one.path.startsWith(path.join(dir,'备份')+path.sep));assert.notEqual(one.path,two.path);assert.match(path.basename(one.path),/^\.\._\.\._方案_v1-\d{8}-\d{6}\.pdf$/);
 assert.equal(fs.readFileSync(one.path,'latin1'),'%PDF-1');assert.equal(fs.readFileSync(two.path,'latin1'),'%PDF-2');
});
test('Office 版本地化工具：与 WPS 版同一套计算，结果写到右侧，统一格式先确认',async()=>{
  const w={};vm.runInNewContext(read('public/locale-engine.js')+'\n'+read('public/office/office-tools.js'),{window:w,Promise});const T=w.OfficeAddinTools;
  const mk=(values)=>{const writes=[],fmts=[];const cell=(i,j)=>({load(){},get values(){return [['']];},set values(v){writes.push([i,j,v[0][0]]);},set numberFormat(v){fmts.push([i,j,v[0][0]]);},set formulas(v){writes.push([i,j,v[0][0]]);}});
    const r={rowCount:values.length,columnCount:values[0].length,values,formulas:values.map(row=>row.map(v=>v??'')),load(){},getCell:cell,getOffsetRange:()=>({getCell:(i,j)=>cell(i,j+10)})};
    return {writes,fmts,api:{Excel:{run:fn=>fn({workbook:{getSelectedRange:()=>r},sync:async()=>{}})}}};};
  const a=mk([[1234.5],['x']]);assert.match(await T.run('excel','rmbUpper',a.api,{}),/1 项结果，1 个无法识别/);assert.equal(JSON.stringify(a.writes),JSON.stringify([[0,10,'壹仟贰佰叁拾肆元伍角']]));
  const z=mk([['ＡＢ１']]);const ask=await T.run('excel','zenHan',z.api,{});assert.match(ask.confirm,/英文字母和数字改成半角/);assert.equal(z.writes.length,0);
  await T.run('excel','zenHan',z.api,{confirmed:true});assert.equal(JSON.stringify(z.writes),JSON.stringify([[0,0,'AB1']]));
  const s=mk([['R8.11.12']]);await T.run('excel','toSeireki',s.api,{});assert.equal(JSON.stringify(s.writes),JSON.stringify([[0,10,46338]]));assert.equal(s.fmts[0][2],'yyyy/m/d');
});
test('Office 版 PowerPoint：全部页统一字体先列出字体再确认；动画和备注给出手动步骤',async()=>{
 const T=tools();const box=(type,font,text=true)=>{const f={name:font,load(){}};return {type,textFrame:{hasText:text,load(){},textRange:{font:f}},f};};
 const s1=[box('Placeholder','黑体'),box('TextBox','微软雅黑'),box('Image','x',false)],s2=[box('TextBox','微软雅黑'),box('TextBox','宋体',false)];
 const slides=[s1,s2].map(items=>({shapes:{items,load(){}}}));const api={PowerPoint:{run:fn=>fn({presentation:{slides:{items:slides,load(){}}},sync:async()=>{}})}};
 const ask=await T.run('powerpoint','unifyDeckFont',api);assert.match(ask.confirm,/2 页共 3 个文字框，用了 2 种字体：微软雅黑 ×2、黑体 ×1/);assert.match(ask.confirm,/统一为“微软雅黑”/);
 assert.match(await T.run('powerpoint','unifyDeckFont',api,{confirmed:true}),/3 个文字框统一为“微软雅黑”/);assert.equal(s1[0].f.name,'微软雅黑');assert.equal(s2[1].f.name,'宋体');
 assert.match(await T.run('powerpoint','unifyDeckFont',api),/已经是“微软雅黑”/);
 assert.match((await T.run('powerpoint','unifyDeckFont',api,{font:'Meiryo'})).confirm,/统一为“Meiryo”/);
 await assert.rejects(T.run('powerpoint','removeAnimations',api),/动画窗格/);await assert.rejects(T.run('powerpoint','removeNotes',api),/检查文档/);
 assert.equal(T.list('powerpoint').find(x=>x.id==='unifyDeckFont').fields[0].key,'font');
});
test('Office 侧栏：taskpane.html 引用的每个文件都能从 https 服务取到，侧栏用到的接口都会转发',async()=>{
  const fsm=await import('node:fs');const html=fsm.readFileSync(new URL('../public/office/taskpane.html',import.meta.url),'utf8');
  const refs=[...html.matchAll(/(?:src|href)="\.\.(\/[^"]+)"/g)].map(m=>m[1]);assert.ok(refs.length>10);
  for(const r of refs)assert.ok(officeFileAllowed(r),'Office 侧栏取不到 '+r);
  const shared=['ai-text-panel.js','ai-chat-panel.js','file-tools-panel.js','quick-panel.js','formula-panel.js','ai-formula.js'].map(f=>fsm.readFileSync(new URL('../public/'+f,import.meta.url),'utf8')).concat(fsm.readdirSync(new URL('../public/office/',import.meta.url)).filter(f=>f.endsWith('.js')).map(f=>fsm.readFileSync(new URL('../public/office/'+f,import.meta.url),'utf8'))).join('\n');
  const used=[...new Set([...shared.matchAll(/['"](\/api\/[a-z/-]+)['"]/g)].map(m=>m[1]))].filter(p=>!/^\/api\/(ai\/models|ai\/test|health)$/.test(p));
  const src=fsm.readFileSync(new URL('../lib/office-server.mjs',import.meta.url),'utf8');for(const p of used)assert.ok(src.includes("'"+p+"'"),'Office 侧栏的接口没有转发：'+p);
});
