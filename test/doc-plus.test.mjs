import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
const code=toolsCode();
function bridge(a,confirm=()=>true){const asked=[];const window={Application:a,confirm:m=>{asked.push(m);return confirm(m);}};vm.runInNewContext(code+'\n'+fs.readFileSync(new URL('../public/postal-format.js',import.meta.url),'utf8')+'\n'+fs.readFileSync(new URL('../public/tools/postal.js',import.meta.url),'utf8'),{window,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date});window.OfficeTools.asked=asked;return window.OfficeTools;}

// 简化的 WPS 文字文档：段落按顺序排，含 \f（分页符）的段落之后换页；sec=true 表示段落里有分节符。
function wordDoc(paras,opts={}){
  const P=paras.map(p=>typeof p==='string'?{text:p}:{...p});
  function layout(){let pos=0,page=1;P.forEach(p=>{p.start=pos;p.end=pos+p.text.length;pos=p.end;p.page=page;if(/[\f]/.test(p.text)||p.sec)page++;});return page;}
  function pageAt(x){layout();const p=P.find(q=>x>=q.start&&x<q.end)||P[P.length-1];return p.page;}
  const shrunk=[];
  function fmt(){return {Font:{},ParagraphFormat:{}};}
  function range(s,e){return {Start:s,End:e,get Text(){layout();let out='';P.forEach(p=>{const a=Math.max(s,p.start),b=Math.min(e,p.end);if(b>a)out+=p.text.slice(a-p.start,b-p.start);});return out;},Information(k){return k===3?pageAt(e>s?e-1:s):0;},Delete(){layout();const hit=P.filter(p=>p.start>=s&&p.end<=e);hit.forEach(p=>P.splice(P.indexOf(p),1));},...own(s,e)};}
  function own(s,e){layout();const p=P.find(q=>q.start===s&&q.end===e);return p?{Font:p.font||(p.font={}),ParagraphFormat:p.pf||(p.pf={})}:fmt();}
  function paraRange(p){const r={get Start(){layout();return p.start;},get End(){layout();return p.end;},get Text(){return p.text;},Information(k){layout();if(k===3)return p.page;if(k===12)return !!p.table;return 0;},InlineShapes:{Count:p.pics||0},Tables:{Count:p.table?1:0},Font:p.font||(p.font={}),ParagraphFormat:p.pf||(p.pf={})};return r;}
  const sections=()=>{layout();const ends=P.filter(p=>p.sec).map(p=>({Range:{End:p.end}}));ends.push({Range:{End:P[P.length-1].end}});return ends;};
  const d={Name:'test.docx',TrackRevisions:false,
    get Content(){return {Information(k){return k===4?layout():0;},get End(){layout();return P[P.length-1].end;}};},
    get Paragraphs(){return {get Count(){return P.length;},Item(i){const p=P[i-1];const r=paraRange(p);if(i===P.length)shrunk.push(r);return {Range:r,OutlineLevel:p.level||10,Style:{NameLocal:p.style||'正文'}};}};},
    get Sections(){const s=sections();return {Count:s.length,Item(i){return s[i-1];}};},
    Shapes:{Count:0},Range:range};
  const a={ActiveDocument:d,Selection:{Text:'',Range:{Start:0,End:0}}};
  return {a,d,P,shrunk};
}

test('带详细确认的文档工具与 FORM_FIRST 工具不再叠加通用确认',()=>{
 const T=bridge({ActiveDocument:{}});
 for(const id of ['deleteBlankPages','officialFormatCn','businessFormatJa','landscapePages','fixPageJump','pictureFix','noBreakLatin','tableBeautify','boldInlineHeads','alignOptions','wordToPpt','jaStyle','proofCheck','captionNumbers','fixListNumbering'])assert.ok(T.describe(id,'wps').ownConfirm.includes('wps'),id);
 for(const id of ['postalFill','wanUnit','idCheck','mergeWorkbooks','codesBatch'])assert.ok(T.describe(id,'et').ownConfirm.includes('et'),id);
 for(const id of ['slidesToImages','handoutPdf'])assert.ok(T.describe(id,'wpp').ownConfirm.includes('wpp'),id);
 assert.ok(!T.describe('removeWatermark','wps').ownConfirm?.includes('wps'),'没有自身确认的工具保留通用确认');
});

test('删除空白页：只删空白页上的空段落和分页符，内容页不动',async()=>{
  const {a,P}=wordDoc(['第一页正文\r','\f\r','\r','\f\r','第三页正文\r']);
  const T=bridge(a);const r=await T.run('deleteBlankPages',{host:'wps'});
  assert.match(T.asked[0],/第 2 页/);
  assert.deepEqual(P.map(p=>p.text),['第一页正文\r','\f\r','第三页正文\r']);
  assert.match(r,/已删除 2 处.*原来共 3 页，空白的是第 2 页/);
});
test('删除空白页：文末两个 Ctrl+Enter（含空格）删掉，保留最终段落标记',async()=>{
 const m=wordDoc(['正文\r',' \f\r','\f\r']);
 const original=m.d.Range;m.d.Range=(s,e)=>{const r=original(s,e);r.Delete=()=>{const p=m.P.find(p=>s>=p.start&&e<=p.end);if(p&&s===p.start&&e===p.end-1)p.text=p.text.slice(-1);else {const hit=m.P.filter(p=>p.start>=s&&p.end<=e);hit.forEach(p=>m.P.splice(m.P.indexOf(p),1));}};return r;};
 const T=bridge(m.a);assert.match(await T.runFromUi('deleteBlankPages'),/已删除 2 处/);
 assert.equal(T.asked.length,1);assert.equal(m.P.map(p=>p.text).join(''),'正文\r\r');assert.equal(m.d.Content.Information(4),1);
});

test('日文商务文书：日期、收件人、发件人和真正标题分别排版；公文日期和发文字号不当标题',async()=>{
 function model(lines){const ps=lines.map(t=>({Range:{Text:t+'\r',Font:{Size:10.5,Bold:0},ParagraphFormat:{},Information:()=>false}}));ps.forEach((p,i)=>{p.Range.Start=i;p.Range.End=i+1;});
 const section={PageSetup:{},Footers:{Item:()=>({Range:{Text:'已有页脚'}})}};
 const d={Name:'t.docx',Content:{Paragraphs:{Count:ps.length,Item:i=>ps[i-1]}},Sections:{Count:1,Item:()=>section},Range:s=>ps[s].Range};return {ps,T:bridge({ActiveDocument:d})};}
 const m=model(['2026年11月26日','株式会社お客様','営業部 御中','株式会社送り主','年末年始休業のお知らせ','拝啓 平素より','記','休業期間','以上']);
 await m.T.runFromUi('businessFormatJa');assert.equal(m.T.asked.length,1);
 assert.deepEqual(m.ps.map(p=>p.Range.ParagraphFormat.Alignment),[2,0,0,2,1,3,1,3,2]);
 assert.equal(m.ps[4].Range.Font.Size,14);assert.equal(m.ps[4].Range.Font.Bold,1);assert.equal(m.ps[0].Range.Font.Bold,0);
 for(const i of [0,1,2,3,4,6,8])assert.equal(m.ps[i].Range.ParagraphFormat.FirstLineIndent,0);
 const cn=model(['2026年11月26日','办公〔2026〕12号','关于放假的通知','这是正文。']);await cn.T.runFromUi('officialFormatCn');
 assert.equal(cn.T.asked.length,1);assert.equal(cn.ps[2].Range.Font.Size,22);assert.equal(cn.ps[0].Range.Font.Size,16);assert.equal(cn.ps[1].Range.Font.Size,16);
});

test('删除空白页：分节符造成的空白页不删除，只提示',async()=>{
  const {a,P}=wordDoc(['正文\r',{text:'\r',sec:true},'下一节正文\r']);
  P[0].text='正文\f\r';
  const T=bridge(a);const r=await T.run('deleteBlankPages',{host:'wps'});
  assert.match(r,/分节符/);assert.equal(P.length,3);assert.equal(T.asked.length,0);
});
test('删除空白页：表格后面的空白页只把文末空段落缩小，不删除',async()=>{
  const {a,P,shrunk}=wordDoc([{text:'表格\r',table:true},{text:'表格\f\r',table:true},'\r']);
  const T=bridge(a);await T.run('deleteBlankPages',{host:'wps'});
  assert.equal(P.length,3);const last=P[2];assert.equal(last.font.Size,1);assert.equal(last.pf.LineSpacing,1);assert.equal(last.pf.LineSpacingRule,4);
});
test('删除空白页：取消时不做任何修改；没有空白页时直接说明',async()=>{
  let m=wordDoc(['一\f\r','\r','\f\r','二\r']);const T=bridge(m.a,()=>false);assert.match(await T.run('deleteBlankPages',{host:'wps'}),/已取消/);assert.equal(m.P.length,4);
  m=wordDoc(['一\f\r','二\r']);assert.match(await bridge(m.a).run('deleteBlankPages',{host:'wps'}),/^没有找到空白页/);
});
test('删除空白页：修订模式下拒绝执行',async()=>{const m=wordDoc(['一\f\r','\r','二\r']);m.d.TrackRevisions=true;await assert.rejects(bridge(m.a).run('deleteBlankPages',{host:'wps'}),/修订/);});

function textDoc(text){let content=text;const d={Name:'t.docx',TrackRevisions:false,Range(s,e){return {InsertAfter(x){content=content.slice(0,e)+x+content.slice(e);},Delete(){content=content.slice(0,s)+content.slice(e);}};}};
  const a={ActiveDocument:d,Selection:{get Text(){return content;},Range:{Start:0,StoryType:1,Fields:{Count:0},get Text(){return content;}}}};return {a,get:()=>content};}
test('中英文之间加空格、去空格（中文和日文）',async()=>{
  let m=textDoc('使用Excel处理2026年的数据，WPS版本12');await bridge(m.a).run('addCjkSpace',{host:'wps'});assert.equal(m.get(),'使用 Excel 处理 2026 年的数据，WPS 版本 12');
  m=textDoc('使用 Excel 处理 2026 年，資料を Word で');await bridge(m.a).run('removeCjkSpace',{host:'wps'});assert.equal(m.get(),'使用Excel处理2026年，資料をWordで');
  m=textDoc('已经 OK 了');assert.match(await bridge(m.a).run('addCjkSpace',{host:'wps'}),/已经有空格/);
});
test('行内标题和选项的识别规则',()=>{
  const T=bridge(wordDoc(['x\r']).a),D=T.DocPlus;
  assert.equal(D.HEAD.exec('1、工作目标：完成年度计划')[1],'1、工作目标：');
  assert.equal(D.HEAD.exec('（三）保障措施：加强领导')[1].trim(),'（三）保障措施：');
  assert.equal(D.HEAD.exec('这句话里有冒号：但不是标题'),null);
  assert.deepEqual([...D.OPT.exec('A. 北京 B. 上海 C. 广州 D. 深圳').slice(1)],['北京','上海','广州','深圳']);
  assert.deepEqual([...D.OPT.exec('A．对 B．错 C．不确定').slice(1,4)],['对','错','不确定']);
});
test('文档整理工具只在 WPS 文字里登记，并归到“文档整理”',()=>{
  const T=bridge(wordDoc(['x\r']).a),ids=[...T.list('wps')].filter(x=>x.group==='docfix').map(x=>x.id);
  assert.deepEqual(ids,['deleteBlankPages','officialFormatCn','businessFormatJa','fixPageJump','landscapePages','tableBeautify','pictureFix','boldInlineHeads','alignOptions','noBreakLatin','removeWatermark','wordToPpt','addCjkSpace','removeCjkSpace']);
  assert.equal(T.list('et').some(x=>x.group==='docfix'),false);
});
test('Word（Office）版与 WPS 版工具一致，并合并进 Word 工具列表',()=>{
  const ctx={};ctx.window=ctx;vm.runInNewContext(fs.readFileSync(new URL('../public/office/office-docplus.js',import.meta.url),'utf8')+'\n'+fs.readFileSync(new URL('../public/office/office-tools.js',import.meta.url),'utf8'),ctx);
  const ids=[...ctx.OfficeDocPlus.list.map(x=>x[0])];const wps=[...bridge(wordDoc(['x\r']).a).list('wps')].filter(x=>x.group==='docfix').map(x=>x.id);
  assert.deepEqual(ids,wps);for(const id of ids)assert.equal(typeof ctx.OfficeDocPlus.tools[id],'function',id);
  const word=ctx.OfficeAddinTools.list('word').map(x=>x.id);for(const id of ids)assert.ok(word.includes(id),id);
  assert.equal(ctx.OfficeAddinTools.list('excel').some(x=>x.id==='deleteBlankPages'),false);
});

// ---- Word（Office.js）模拟 ----
function officeWord(paras){
  const P=paras.map(p=>typeof p==='string'?{text:p,ox:''}:p),deleted=[];
  const items=()=>P.map((p,i)=>({text:p.text,tableNestingLevel:p.table?1:0,font:{},getOoxml(){return {value:p.ox||'<w:p><w:r><w:t>'+p.text+'</w:t></w:r></w:p>'};},delete(){deleted.push(i);}}));
  let cache=null;
  const context={document:{body:{get paragraphs(){cache=items();return {items:cache,load(){}};}},getSelection(){return {text:'',load(){}};}},sync:async()=>{}};
  return {api:{Word:{run:fn=>fn(context)}},deleted,last:()=>cache[cache.length-1]};
}
test('Word 版删除空白页：文末空段落和连续分页符之间的空段落，先确认',async()=>{
  const ctx={};ctx.window=ctx;vm.runInNewContext(fs.readFileSync(new URL('../public/office/office-docplus.js',import.meta.url),'utf8'),ctx);const D=ctx.OfficeDocPlus.tools;
  const br='<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  let m=officeWord(['正文',{text:'',ox:br},'',{text:'',ox:br},'第三页','','']);
  const first=await D.deleteBlankPages(m.api,{});assert.match(first.confirm,/将删除 3 个/);assert.deepEqual(m.deleted,[]);
  await D.deleteBlankPages(m.api,{confirmed:true});assert.deepEqual([...m.deleted].sort(),[1,2,5]);
  m=officeWord([{text:'表',table:true},'']);await D.deleteBlankPages(m.api,{confirmed:true});assert.deepEqual(m.deleted,[]);assert.equal(m.last().font.size,1);
  m=officeWord(['只有正文']);assert.match(await D.deleteBlankPages(m.api,{}),/^没有找到/);
  m=officeWord(['正文',{text:'',ox:'<w:p><w:r><w:drawing/></w:r></w:p>'},'']);assert.match(await D.deleteBlankPages(m.api,{}),/^没有找到/);
});
test('Word 版不能改横向时给出手动步骤',async()=>{
  const ctx={};ctx.window=ctx;vm.runInNewContext(fs.readFileSync(new URL('../public/office/office-docplus.js',import.meta.url),'utf8'),ctx);
  await assert.rejects(ctx.OfficeDocPlus.tools.landscapePages(),/纸张方向/);
});

test('Word 转 PPT：按标题整理大纲（跳过表格里的段落），确认后交给版式工作室生成并打开',async()=>{
  const {a}=wordDoc([{text:'季度汇报\r',style:'标题'},{text:'一、成果\r',level:1,style:'标题 1'},{text:'销售额增长 15%。\r'},{text:'表格里的字\r',table:true},{text:'二、计划\r',level:1,style:'标题 1'},{text:'扩大市场\r'}]);
  const calls=[];const fetch=async(url,o)=>{calls.push({url,body:o&&o.body&&JSON.parse(o.body)});return url==='/api/session'?{ok:true,json:async()=>({token:'t'})}:{ok:true,json:async()=>({opened:true,pages:5})};};
  const asked=[];const window={Application:a,confirm:m=>{asked.push(m);return true;}};vm.runInNewContext(code,{window,fetch,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date,setTimeout,clearTimeout});
  const r=await window.OfficeTools.run('wordToPpt',{host:'wps'});
  assert.match(asked[0],/整理出 2 节：“成果”、“计划”/);assert.match(r,/已生成 5 页 PPT/);
  const body=calls[1].body;assert.equal(calls[1].url,'/api/studio/open');assert.equal(body.title,'季度汇报');assert.equal(body.outline,'成果\n- 销售额增长 15%。\n\n计划\n- 扩大市场');assert.equal(body.theme,'midnight');
  const no={Application:wordDoc(['只有正文。\r']).a,confirm:()=>true};vm.runInNewContext(code,{window:no,fetch,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date,setTimeout,clearTimeout});
  await assert.rejects(no.OfficeTools.run('wordToPpt',{host:'wps'}),/没有找到标题/);
});
