import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
function load(a,{exists}={}){const asked=[];const window={Application:a,confirm:m=>{asked.push(m);return true;}};
  const fetch=exists?async(url,o)=>({ok:true,json:async()=>url==='/api/session'?{token:'t'}:{exists:JSON.parse(o.body).paths.map(exists)}}):async()=>{throw Error('offline');};
  vm.runInNewContext(toolsCode(),{window,fetch,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date,setTimeout,clearTimeout});return {T:window.OfficeTools,B:window.OfficeBatch,asked};}
function dialogs(...answers){let k=0;return ()=>{const ans=answers[k++];return {Title:'',AllowMultiSelect:false,Filters:{Clear(){},Add(){}},InitialFileName:'',Show:()=>ans?-1:0,SelectedItems:{get Count(){return [].concat(ans||[]).length;},Item:i=>[].concat(ans)[i-1]}};};}
// 文档模拟：段落列表 [{text,level}]；Range 取文字；Documents.Add 新建的文档记录 FormattedText 和另存路径。
function wordApp(paras,files={}){
  const text=paras.map(p=>p.text+'\r').join(''),starts=[];let p=0;for(const x of paras){starts.push(p);p+=x.text.length+1;}
  const range=(s,e)=>({Start:s,End:e,get Text(){return text.slice(s,e);},get FormattedText(){return {from:s,to:e,text:text.slice(s,e)};}});
  const made=[],opened=[];
  const doc={Name:'长文.docx',FullName:'C:\\资料\\长文.docx',Path:'C:\\资料',Paragraphs:{Count:paras.length,Item:i=>({OutlineLevel:paras[i-1].level||10,Range:range(starts[i-1],starts[i-1]+paras[i-1].text.length+1)})},Content:{End:text.length},Range:range,PageSetup:{Orientation:0,TopMargin:72,BottomMargin:72,LeftMargin:90,RightMargin:90,PageWidth:595,PageHeight:842},Activate(){}};
  const a={ActiveDocument:doc,made,opened,
    Documents:{Count:1,Item:()=>doc,Add(){const nd={Paragraphs:{Count:1,Item:()=>({Range:{Text:'x'}})},PageSetup:{},Content:{set FormattedText(v){nd.content=v;}},SaveAs2(path,fmt){nd.saved=path;nd.fmt=fmt;},Close(){nd.closed=true;}};made.push(nd);return nd;},
      Open(path,c,ro){const f=files[path];if(!f)throw Error('打不开');const d={Name:path.split('\\').pop(),readOnly:ro,ExportAsFixedFormat(out,fmt){d.pdf=out;d.fmt=fmt;},Close(x){d.closed=x;}};opened.push(d);return d;}}};
  return a;
}
test('按标题拆分：前言单独一份，每个标题 1 一份，文件名按顺序编号，原文档不改',async()=>{
  const a=wordApp([{text:'封面说明'},{text:'第一章 总则',level:1},{text:'内容一'},{text:'1.1 小节',level:2},{text:'第二章 附则',level:1},{text:'内容二'}]);
  a.FileDialog=dialogs(['C:\\输出']);const {T,asked}=load(a,{exists:()=>false});
  const r=await T.run('splitByHeading',{host:'wps'});assert.match(r,/已拆分出 3 个文件/);assert.match(asked[0],/01-前言\.docx\n02-第一章 总则\.docx\n03-第二章 附则\.docx/);
  assert.deepEqual(a.made.map(d=>d.saved),['C:\\输出\\01-前言.docx','C:\\输出\\02-第一章 总则.docx','C:\\输出\\03-第二章 附则.docx']);
  assert.equal(a.made[1].content.text,'第一章 总则\r内容一\r1.1 小节\r');assert.equal(a.made[2].content.text,'第二章 附则\r内容二\r');assert.ok(a.made.every(d=>d.closed&&d.fmt===16));assert.equal(a.made[0].PageSetup.LeftMargin,90);
  const b=wordApp([{text:'第一章',level:1},{text:'甲'},{text:'第二章',level:1},{text:'乙'}]);b.FileDialog=dialogs(['C:\\输出']);
  assert.deepEqual([...load(b).B.headingParts(b.ActiveDocument,1)].map(p=>p.title),['第一章','第二章']);
  const c=wordApp([{text:'没有标题'}]);c.FileDialog=dialogs(['C:\\输出']);await assert.rejects(load(c).T.run('splitByHeading',{host:'wps'}),/没有找到“标题 1”/);
});
test('批量转 PDF：只读打开、导出后关闭；已有同名 PDF 时加序号；打不开的文件单独列出',async()=>{
  const a=wordApp([{text:'x'}],{'D:\\甲.docx':1,'D:\\乙.docx':1});a.FileDialog=dialogs(['D:\\乙.docx','D:\\甲.docx','D:\\丙.docx'],['D:\\PDF']);
  const {T,asked}=load(a,{exists:p=>p==='D:\\PDF\\甲.pdf'});const r=await T.run('batchPdf',{host:'wps'});
  assert.match(asked[0],/3 个文字文档转成 PDF/);assert.deepEqual(a.opened.map(d=>[d.Name,d.readOnly,d.pdf,d.fmt,d.closed]),[['甲.docx',true,'D:\\PDF\\甲-2.pdf',17,0],['乙.docx',true,'D:\\PDF\\乙.pdf',17,0]]);
  assert.match(r,/已转成 PDF 2 个文件/);assert.match(r,/1 个没能处理：丙\.docx（打不开）/);assert.match(r,/原文件都没有改动/);
  const b=wordApp([{text:'x'}],{'D:\\甲.docx':1});b.FileDialog=dialogs(['D:\\甲.docx'],['D:\\PDF']);await assert.rejects(load(b).T.run('batchPdf',{host:'wps'}),/无法确认输出文件/);assert.equal(b.opened.length,0);
  const c=wordApp([{text:'x'}]);c.FileDialog=dialogs(null);assert.equal(await load(c).T.run('batchPdf',{host:'wps'}),'已取消。');
});
test('发送到另一个文档：多个文档时先列出序号；文字连格式插到目标末尾',async()=>{
  const src={FullName:'C:\\a.docx',Name:'a.docx'},tA={FullName:'C:\\b.docx',Name:'b.docx'},tB={FullName:'C:\\c.docx',Name:'c.docx'};
  let put=null;tB.Content={End:11};tB.Range=(s,e)=>({get Text(){return s===9?'\r':'';},InsertParagraphAfter(){},set FormattedText(v){put={at:s,v};}});
  const a={ActiveDocument:src,Documents:{Count:3,Item:i=>[src,tA,tB][i-1]},Selection:{Start:2,End:6,Range:{FormattedText:'选中内容'}}};const {T}=load(a);
  assert.match(await T.run('sendToDoc',{host:'wps'}),/有 2 个，请在“发送到第几个”里填序号：\n1\. b\.docx\n2\. c\.docx/);
  assert.match(await T.run('sendToDoc',{host:'wps',target:'2'}),/已发送到《c\.docx》/);assert.deepEqual(put,{at:10,v:'选中内容'});
});

test('批量输出：多个同名、大小写和已有序号不冲突',async()=>{
 const {B}=load({}, {exists:p=>/报告(?:-2|-3)?\.pdf$/i.test(p)});
 assert.deepEqual(Array.from(await B.freeNames(Array(5).fill('D:/报告.pdf'))),[4,5,6,7,8].map(n=>'D:/报告-'+n+'.pdf'));
 const out=Array.from(await load({}, {exists:()=>false}).B.freeNames(['D:/A.pdf','D:/a.pdf','D:/A-2.pdf']));
 assert.equal(new Set(out.map(p=>p.toLowerCase())).size,3);
});
