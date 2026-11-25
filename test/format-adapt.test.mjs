// 0.3.1：插入/整理时跟随目标格式。替身对象只模拟被调用的属性，真机效果仍需在 WPS 中验收。
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
const code=toolsCode();
function bridge(a,host='wps',fetchImpl){const window={Application:a,confirm:()=>true};vm.runInNewContext(code,{window,fetch:fetchImpl,location:{pathname:'/addon/'+host+'/',search:''},URLSearchParams,Date});return window.OfficeTools;}
const frame=(text,size,name='宋体',extra={})=>({Name:text,HasTextFrame:true,TextFrame:{HasText:true,TextRange:{Text:text,Font:{Name:name,Size:size}}},...extra});
const shapes=list=>({Count:list.length,Item:i=>list[i-1]});

test('统一本页字体保留标题大小，正文统一为最常见字号',async()=>{
  const title=frame('标题',40,'黑体',{PlaceholderFormat:{Type:1}}),a=frame('甲',18),b=frame('乙',18),c=frame('丙',24);
  const slide={Shapes:shapes([title,a,b,c])};
  const app={ActivePresentation:{Name:'d.pptx',Slides:{Count:1,Item:()=>slide}},ActiveWindow:{View:{Slide:slide}}};
  const msg=await bridge(app,'wpp').run('formatSlide',{host:'wpp'});
  assert.equal(title.TextFrame.TextRange.Font.Size,40);
  for(const f of [a,b,c])assert.equal(f.TextFrame.TextRange.Font.Size,18);
  for(const f of [title,a,b,c])assert.equal(f.TextFrame.TextRange.Font.Name,'黑体');
  assert.match(msg,/标题保持原有大小/);
});

test('PPT 插入放在当前页之后，并沿用演示里已有的字体',async()=>{
  const indexes=[],existing={Shapes:shapes([frame('原有',20,'思源黑体')])};
  const d={Name:'d.pptx',PageSetup:{SlideWidth:960,SlideHeight:540},Slides:{Count:5,Item:()=>existing,Add(i){indexes.push(i);this.Count++;const made=[];return {Shapes:{AddTextbox(){const f={TextFrame:{TextRange:{Font:{}}}};made.push(f);return f;}},made,Delete(){}};}}};
  const app={ActivePresentation:d,ActiveWindow:{View:{Slide:{SlideIndex:2}}}};
  const msg=await bridge(app,'wpp').run('makeSlides',{host:'wpp',text:'一\n内容\n\n二\n内容'});
  assert.deepEqual(indexes,[3,4]);
  assert.match(msg,/当前页之后/);
});

test('PPT 表格行列多时自动缩小字号，表头加粗',async()=>{
  const cells=[];const d={Name:'d.pptx',PageSetup:{SlideWidth:960,SlideHeight:540},Slides:{Count:0,Item:()=>({Shapes:shapes([])}),Add(){this.Count++;return {Shapes:{AddTable(r,c){return {Table:{Cell(){const t={Text:'',Font:{}};cells.push(t);return {Shape:{TextFrame:{TextRange:t}}};}}};}},Delete(){}};}}};
  const rows=Array.from({length:10},(_,i)=>Array.from({length:10},(_,j)=>i+'-'+j));
  await bridge({ActivePresentation:d},'wpp').run('insertTable',{host:'wpp',text:rows.map(r=>r.join('\t')).join('\n')});
  assert.ok(cells.every(c=>c.Font.Size>=10&&c.Font.Size<18),'字号应在 10–17 之间');
  assert.equal(cells[0].Font.Bold,-1);
});

test('Word 匹配周围格式：紧挨标题时参照后面的正文，表格适应页宽',async()=>{
  let fitted=0;const heading={Font:{Name:'黑体',Size:22,Color:0},Paragraphs:{Item:()=>({OutlineLevel:1,Range:{Style:'标题 1',ParagraphFormat:{}}})}};
  const body={Font:{Name:'Calibri',NameFarEast:'仿宋',Size:16,Color:0,Bold:0,Italic:0},Paragraphs:{Item:()=>({OutlineLevel:10,Range:{Style:'正文',ParagraphFormat:{LineSpacingRule:1,LineSpacing:24,SpaceAfter:6}}})}};
  const range={Start:10,End:20,Text:'粘进来的文字',Font:{},ParagraphFormat:{},Tables:{Count:1,Item:()=>({AutoFitBehavior(m){assert.equal(m,2);fitted++;}})}};
  const d={Name:'a.docx',FullName:'C:/a.docx',Content:{End:100},Range:(s)=>s===9?heading:body};
  const msg=await bridge({ActiveDocument:d,Selection:{Text:range.Text,Range:range}}).run('matchFormat');
  assert.equal(range.Style,'正文');assert.equal(range.Font.Name,'Calibri');assert.equal(range.Font.NameFarEast,'仿宋');assert.equal(range.Font.Bold,0);assert.equal(range.Font.Size,16);assert.equal(fitted,1);
  assert.match(msg,/仿宋/);
});

test('Word 匹配格式遇到混合格式时不写入无效值',async()=>{
  const ref={Font:{Name:'',Size:9999999,Color:9999999},Paragraphs:{Item:()=>({OutlineLevel:10,Range:{Style:'正文',ParagraphFormat:{}}})}};
  const range={Start:5,End:8,Text:'abc',Font:{Name:'原字体',Size:12},ParagraphFormat:{},Tables:{Count:0}};
  await bridge({ActiveDocument:{Name:'a.docx',Content:{End:50},Range:()=>ref},Selection:{Text:'abc',Range:range}}).run('matchFormat');
  assert.equal(range.Font.Name,'原字体');assert.equal(range.Font.Size,12);
});

test('PPT 匹配本页格式：跟随正文字体，超出页面的对象缩放回来',async()=>{
  const body=frame('正文',20,'思源黑体'),pasted=frame('粘贴',36,'Arial',{Width:1200,Height:300,Left:-10,Top:20});
  const slide={Shapes:shapes([body,pasted])};
  const app={ActivePresentation:{Name:'d.pptx',PageSetup:{SlideWidth:960,SlideHeight:540},Slides:{Count:1,Item:()=>slide}},ActiveWindow:{View:{Slide:slide},Selection:{Type:2,ShapeRange:shapes([pasted])}}};
  const msg=await bridge(app,'wpp').run('matchFormat',{host:'wpp'});
  assert.equal(pasted.TextFrame.TextRange.Font.Name,'思源黑体');assert.equal(pasted.TextFrame.TextRange.Font.Size,20);
  assert.ok(pasted.Width<=900&&pasted.Left>=0);assert.match(msg,/缩放进页面/);
});

test('表格中匹配格式会提示使用格式刷，而不是静默失败',async()=>{
  const app={ActiveWorkbook:{Name:'t.xlsx'},ActiveSheet:{Name:'S'},Selection:{Rows:{Count:1},Columns:{Count:1}}};
  await assert.rejects(bridge(app,'et').run('matchFormat',{host:'et'}),/格式刷/);
});

test('美化表头使用浅色底深色字，打印也清楚',async()=>{
  const header={Font:{},Interior:{},Borders:{Item:()=>({})}};
  const app={ActiveWorkbook:{Name:'t.xlsx'},ActiveSheet:{Name:'S'},Selection:{Rows:{Count:3,Item:()=>header},Columns:{Count:2}}};
  await bridge(app,'et').run('formatHeader',{host:'et'});
  assert.equal(header.Interior.Color,221+235*256+247*65536);assert.equal(header.Font.Bold,true);
});

test('复制当前页：副本被放到别处时移回原页之后并选中',async()=>{
  const moves=[];let selected=false;const copy={SlideIndex:4,MoveTo(i){moves.push(i);this.SlideIndex=i;},Select(){selected=true;}};
  const d={Name:'d.pptx',Slides:{Count:4,Range:i=>{assert.equal(i,2);return {Duplicate:()=>({Item:()=>copy})};}}};
  const msg=await bridge({ActivePresentation:d,ActiveWindow:{View:{Slide:{SlideIndex:2}}}},'wpp').run('duplicateSlide',{host:'wpp'});
  assert.deepEqual(moves,[3]);assert.ok(selected);assert.match(msg,/第 2 页之后/);
});

test('新建演示还没点过页面时，从选中页或第 1 页判断当前页',async()=>{
  let ranged;const only={SlideIndex:1};
  const d={Name:'d.pptx',Slides:{Count:1,Item:()=>only,Range:i=>{ranged=i;return {Duplicate:()=>null};}}};
  await bridge({ActivePresentation:d,ActiveWindow:{View:{},Selection:{}}},'wpp').run('duplicateSlide',{host:'wpp'});
  assert.equal(ranged,1);
});
