import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
// 简化的 WPS 演示：每页有若干文字框（含表格、组合）、动画序列、切换效果和备注。
function deck(spec){
  const tr=(text,font)=>({Text:text,Font:{Name:font,NameFarEast:font}});
  const shape=x=>x.table?{HasTable:true,Table:{Rows:{Count:1},Columns:{Count:x.table.length},Cell:(r,c)=>({Shape:{TextFrame:{HasText:true,TextRange:x.table[c-1]}}})}}:x.group?{Type:6,GroupItems:{Count:x.group.length,Item:i=>shape(x.group[i-1])}}:{HasTextFrame:true,TextFrame:{HasText:!!x.text,TextRange:x.range||(x.range=tr(x.text,x.font))}};
  const slides=spec.map(s=>{const seq=[...Array(s.anims||0)].map(()=>({Delete(){seq.splice(seq.indexOf(this),1);}}));const shapes=s.shapes.map(shape);const notes=tr(s.notes||'','x');
    return {seq,notes,shapes,trans:{EntryEffect:s.trans||0},obj:{Shapes:{Count:shapes.length,Item:i=>shapes[i-1]},TimeLine:{MainSequence:{get Count(){return seq.length;},Item:i=>seq[i-1]},InteractiveSequences:{Count:0}},get SlideShowTransition(){return this._t;},_t:null,NotesPage:{Shapes:{Count:2,Item:i=>i===1?{PlaceholderFormat:{Type:1}}:{PlaceholderFormat:{Type:2},HasTextFrame:true,TextFrame:{TextRange:notes}}}}}};});
  slides.forEach(s=>{s.obj._t=s.trans;});
  const d={Slides:{Count:slides.length,Item:i=>slides[i-1].obj}};return {a:{ActivePresentation:d,ActiveWindow:{}},slides};
}
function load(a,answer=true){const asked=[];const window={Application:a,confirm:m=>{asked.push(m);return answer;}};vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/wpp/',search:''},URLSearchParams,Date});return {T:window.OfficeTools,asked};}
test('全部页统一字体：文字框、表格、组合都改，默认用最常见的字体',async()=>{
  const cellA={Text:'甲',Font:{Name:'宋体',NameFarEast:'宋体'}},cellB={Text:'乙',Font:{Name:'微软雅黑',NameFarEast:'微软雅黑'}};
  const {a,slides}=deck([{shapes:[{text:'标题',font:'黑体'},{text:'正文',font:'微软雅黑'}]},{shapes:[{table:[cellA,cellB]},{group:[{text:'组合里',font:'微软雅黑'}]}]}]);
  const {T,asked}=load(a);const r=await T.run('unifyDeckFont',{host:'wpp'});
  assert.match(asked[0],/全部 2 页共 5 处文字，现在用了 3 种字体：微软雅黑 ×3/);assert.match(r,/已把 5 处文字统一为“微软雅黑”/);
  assert.equal(slides[0].shapes[0].TextFrame.TextRange.Font.NameFarEast,'微软雅黑');assert.equal(cellA.Font.Name,'微软雅黑');
  assert.match(await T.run('unifyDeckFont',{host:'wpp'}),/不需要修改/);
  assert.match(await load(a).T.run('unifyDeckFont',{host:'wpp',font:'Meiryo'}),/统一为“Meiryo”/);
});
test('删除全部动画和切换效果、删除全部备注：先报数量，删完再核对',async()=>{
  const {a,slides}=deck([{shapes:[{text:'a',font:'x'}],anims:3,trans:257,notes:'讲稿一'},{shapes:[],anims:0},{shapes:[],anims:2,notes:'  '}]);
  let {T,asked}=load(a);assert.match(await T.run('removeAnimations',{host:'wpp'}),/已删除 5 个动画和 1 页的切换效果/);assert.match(asked[0],/2 页上的 5 个动画，以及 1 页的切换效果/);
  assert.equal(slides[0].seq.length+slides[2].seq.length,0);assert.equal(slides[0].trans.EntryEffect,0);
  assert.match(await T.run('removeAnimations',{host:'wpp'}),/没有找到动画/);
  ({T,asked}=load(a));assert.match(await T.run('removeNotes',{host:'wpp'}),/已清空 1 页的备注/);assert.match(asked[0],/第 1 页/);assert.equal(slides[0].notes.Text,'');
  assert.match(await T.run('removeNotes',{host:'wpp'}),/没有找到备注/);
  const kept=deck([{shapes:[],anims:1}]);assert.match(await load(kept.a,false).T.run('removeAnimations',{host:'wpp'}),/已取消/);assert.equal(kept.slides[0].seq.length,1);
});
test('美化已有演示：统一各页标题以第一个内容页为准，封面不动；加页码时封面不显示',async()=>{
  const mk=(type,text,l,size)=>({PlaceholderFormat:{Type:type},HasTextFrame:true,Left:l,Top:20,Width:600,Height:60,TextFrame:{TextRange:{Text:text,Font:{Name:'宋体',NameFarEast:'宋体',Size:size,Bold:true,Color:{RGB:255}},ParagraphFormat:{Alignment:1}}}});
  const slides=[[mk(3,'封面',100,44)],[mk(1,'第一页',40,32)],[mk(1,'第二页',80,28)],[mk(1,'第三页',60,36)]].map(list=>({Shapes:{Count:list.length,Item:i=>list[i-1]},HeadersFooters:{SlideNumber:{Visible:false}}}));
  const d={Slides:{Count:4,Item:i=>slides[i-1]},PageSetup:{SlideWidth:960,SlideHeight:540}};const a={ActivePresentation:d};
  const {vm,toolsCode}={vm:(await import('node:vm')).default,toolsCode:(await import('./helpers/tools-code.mjs')).toolsCode};
  const window={Application:a,confirm:()=>true};vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/wpp/',search:''},URLSearchParams,Date});const T=window.OfficeTools;
  assert.match(await T.run('unifyTitles',{host:'wpp'}),/已统一 3 页的标题/);
  assert.deepEqual(slides.map(s=>[s.Shapes.Item(1).Left,s.Shapes.Item(1).TextFrame.TextRange.Font.Size]),[[100,44],[40,32],[40,32],[40,32]]);
  assert.match(await T.run('slideNumbers',{host:'wpp'}),/已给 3 页打开页码，封面不显示/);assert.deepEqual(slides.map(s=>s.HeadersFooters.SlideNumber.Visible),[false,true,true,true]);
});
test('图片图表自动排版：按原比例排成网格、间距一致；一段文字时左文右图',async()=>{
  const pic=(w,h,l,t,type=13)=>({Type:type,Width:w,Height:h,Left:l,Top:t,LockAspectRatio:-1});
  const title={PlaceholderFormat:{Type:1},HasTextFrame:true,TextFrame:{TextRange:{Text:'标题'}},Top:20,Height:60,Left:40,Width:880};
  const shapes=[title,pic(400,300,500,300),pic(200,150,10,100),pic(300,300,600,90,3)];
  const slide={SlideIndex:2,Shapes:{get Count(){return shapes.length;},Item:i=>shapes[i-1]}};
  const d={PageSetup:{SlideWidth:960,SlideHeight:540},Slides:{Count:1,Item:()=>slide}};
  const window={Application:{ActivePresentation:d,ActiveWindow:{View:{Slide:slide}}},confirm:()=>true};
  vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/wpp/',search:''},URLSearchParams,Date});
  const r=await window.OfficeTools.run('arrangeVisuals',{host:'wpp'});assert.match(r,/已排好第 2 页的 3 张图片 \/ 图表（1 行 × 3 列）/);
  const v=shapes.slice(1);const tops=v.map(x=>x.Top),lefts=v.map(x=>x.Left);
  assert.ok(Math.abs(v[0].Width/v[0].Height-400/300)<1e-6&&Math.abs(v[2].Width/v[2].Height-1)<1e-6);assert.ok(v.every(x=>x.Top>=20+60&&x.Top+x.Height<=540*0.94+0.01&&x.Left>=48-0.01&&x.Left+x.Width<=912+0.01));
  assert.ok(lefts[1]<lefts[2]&&lefts[2]<lefts[0],'阅读顺序按原来的位置：上面的在前');assert.equal(v.every(x=>x.LockAspectRatio===-1),true);
  const text={HasTextFrame:true,TextFrame:{TextRange:{Text:'要点'}},Top:100,Height:80,Left:40,Width:500};shapes.push(text);
  assert.match(await window.OfficeTools.run('arrangeVisuals',{host:'wpp'}),/左边文字、右边图片/);assert.equal(text.Width,960*0.42);assert.ok(v.every(x=>x.Left>=480-0.01));
  shapes.splice(1,3);await assert.rejects(window.OfficeTools.run('arrangeVisuals',{host:'wpp'}),/没有图片或图表/);
});
