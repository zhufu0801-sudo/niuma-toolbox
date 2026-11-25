import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {exportStudio,planStudio,themes,chooseLayout} from '../lib/studio.mjs';
import {addMotionToSlide,timingXml,transitionXml} from '../lib/pptx-motion.mjs';
const deck={title:'项目阶段工作汇报',subtitle:'第三季度',outline:'阶段成果\n完成第一版\n收集反馈\n建立测试\n\n关键数据\n35% 效率提升\n1200 位用户\n\n推进计划\n2026年1月：需求确认\n2026年3月：试点\n2026年6月：推广\n\n总结\n把复杂的事情做简单，把简单的事情做到极致。'};
test('六套模板版式各不相同：封面、章节页、卡片和编号方式都有区别',()=>{
  const baseThemes=['midnight','neon','swiss','forest','wa','coral'];
  assert.ok(baseThemes.every(id=>themes[id]));
  const sig=id=>{const p=planStudio({...deck,theme:id});return JSON.stringify(p.pages.slice(0,4).map(pg=>pg.elements.map(e=>[e.type,Math.round(e.x/40),Math.round(e.y/40)])));};
  const all=baseThemes.map(sig);assert.equal(new Set(all).size,all.length);
  assert.equal(planStudio({...deck,theme:'business'}).theme,'midnight');
  const p=planStudio({...deck,theme:'swiss'});assert.deepEqual(p.pages.map(x=>x.role),['cover','agenda','divider','content','divider','content','divider','content','divider','content','closing']);
  assert.deepEqual(p.pages.filter(x=>x.role==='content').map(x=>x.layout),['grid','stat','timeline','quote']);
  assert.equal(chooseLayout({points:['a','b'],layout:'timeline'}),'timeline');
  assert.equal(planStudio({...deck,theme:'neon',closing:false}).pages.at(-1).role,'content');
});
test('切换与进入动画：写在 </p:sld> 之前，只引用页面上真实存在的形状，自动播放',()=>{
  const xml='<p:sld><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="e1"/></p:nvSpPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="e2"/></p:nvSpPr></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>';
  const out=addMotionToSlide(xml,{transition:{type:'push',dir:'u'},items:[{name:'e1',effect:'fade'},{name:'e2',effect:'rise',delay:200},{name:'missing',effect:'zoom'}]});
  assert.ok(out.indexOf('<p:transition')>out.indexOf('</p:clrMapOvr>')&&out.indexOf('<p:timing>')>out.indexOf('<p:transition')&&out.endsWith('</p:timing></p:sld>'));
  assert.match(out,/<p:push dir="u"\/>/);assert.equal((out.match(/presetClass="entr"/g)||[]).length,2);assert.match(out,/spid="2"/);assert.match(out,/spid="3"/);assert.ok(!out.includes('spid=""'));
  assert.match(out,/evt="onBegin"/);assert.match(out,/nodeType="afterEffect"/);assert.match(out,/nodeType="withEffect"/);
  assert.equal(addMotionToSlide(out,{transition:{type:'fade'}}),out);assert.equal(transitionXml({type:'origami'}),'');assert.equal(timingXml([]),'');
});
for(const id of Object.keys(themes))test(`${id}：导出的 PPTX 每页都有切换效果，内容页有进入动画，并能用日文排版`,async()=>{
  const {buffer,plan}=await exportStudio({...deck,theme:id});const zip=await JSZip.loadAsync(buffer);
  for(let i=1;i<=plan.pages.length;i++){const x=await zip.file(`ppt/slides/slide${i}.xml`).async('string');assert.match(x,/<p:transition /);
    const ids=new Set([...x.matchAll(/<p:cNvPr id="(\d+)"/g)].map(m=>m[1]));for(const m of x.matchAll(/spid="(\d+)"/g))assert.ok(ids.has(m[1]),`slide${i} 引用了不存在的形状 ${m[1]}`);
    if(plan.pages[i-1].role==='content')assert.match(x,/presetClass="entr"/);}
  const ja=planStudio({title:'新製品のご紹介',outline:'特長\n速い：処理が2倍\n軽い：重量30%減\n静か：騒音を半分に',theme:id});assert.equal(ja.lang,'ja');assert.ok(ja.pages.some(p=>p.elements.some(e=>/Yu /.test(e.face))));
});
import fsx from 'node:fs';
const fix=n=>fsx.readFileSync(new URL('./fixtures/'+n,import.meta.url),'utf8').trim();
const LOGO='data:image/png;base64,'+fix('logo-400x200.png.b64'),PHOTO='data:image/jpeg;base64,'+fix('photo-300x600.jpg.b64');
const rich={title:'季度经营分析',outline:'方案对比\n现行方案｜新方案\n人工录入\n报表分散\n自动采集\n统一看板\n> 强调汇总时间从一周缩短到一天。\n\n各区域销售额\n华东：320 万\n华南：260 万\n华北：190 万\n\n门店实景 [图文]\n新门店 9 月开业\n日均客流 1200 人次\n\n下一步\n上线\n培训\n复盘'};
test('新版式：对比、图表、图文自动识别；演讲备注写在对应页',()=>{
  const p=planStudio({...rich,theme:'midnight'}),content=p.pages.filter(x=>x.role==='content');
  assert.deepEqual(content.map(x=>x.layout),['compare','chart','image','grid']);
  assert.equal(content[0].notes,'强调汇总时间从一周缩短到一天。');assert.ok(!content[0].elements.some(e=>/强调汇总/.test(e.text||'')));
  const ch=content[1].elements.find(e=>e.type==='chart');assert.deepEqual([...ch.labels],['华东','华南','华北']);assert.deepEqual([...ch.values],[320,260,190]);assert.equal(ch.unit,'万');
  assert.ok(content[2].elements.some(e=>/这里放图片/.test(e.text||'')));
  const withImg=planStudio({...rich,theme:'midnight',images:[PHOTO]}).pages.filter(x=>x.layout==='image')[0].elements.find(e=>e.type==='image');
  assert.ok(Math.abs(withImg.w/withImg.h-0.5)<0.01&&withImg.h<=456.5);
});
test('企业品牌：主色替换强调色、深色底用主色深色版，Logo 等比放在封面和页脚；格式不对时给出提示',()=>{
  const p=planStudio({...rich,theme:'midnight',brand:{color:'#0B5FFF',logo:LOGO}});
  const cover=p.pages[0],logo=cover.elements.find(e=>e.type==='image');assert.ok(logo&&Math.abs(logo.w/logo.h-2)<0.01);
  assert.notEqual(cover.background,'1E2761');assert.ok(p.pages.some(pg=>pg.elements.some(e=>e.fill==='0B5FFF'||e.color==='0B5FFF')));
  const content=p.pages.find(x=>x.role==='content');assert.ok(content.elements.some(e=>e.type==='image'&&e.y>=660));
  const light=planStudio({...rich,theme:'swiss',brand:{color:'#FFE600'}});assert.ok(!light.pages.some(pg=>pg.elements.some(e=>e.fill==='FFE600')),'太浅的颜色应自动加深');
  assert.throws(()=>planStudio({...rich,theme:'midnight',brand:{color:'blue'}}),/#RRGGBB/);
  assert.throws(()=>planStudio({...rich,theme:'midnight',images:['data:image/gif;base64,R0lGOD']}),/PNG 或 JPG/);
});
test('导出：图片、原生图表、演讲备注都写进 PPTX',async()=>{
  const {buffer}=await exportStudio({...rich,theme:'coral',brand:{color:'#0B5FFF',logo:LOGO},images:[PHOTO]});const zip=await JSZip.loadAsync(buffer);
  const names=Object.keys(zip.files);assert.ok(names.some(n=>/^ppt\/charts\/chart\d+\.xml$/.test(n)),'应有原生图表');assert.ok(names.filter(n=>/^ppt\/media\//.test(n)).length>=2,'应有 Logo 和图片');
  const notes=await Promise.all(names.filter(n=>/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n)).map(n=>zip.file(n).async('string')));assert.ok(notes.some(x=>x.includes('强调汇总时间从一周缩短到一天')));
});
