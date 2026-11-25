import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
const read=f=>fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8');
function wpsTools(){const window={Application:{},confirm:()=>true};vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date});return window.OfficeTools;}

test('方法卡的“一键完成”都指向已登记的 WPS 文字工具',()=>{
  const ctx={};ctx.window=ctx;vm.runInNewContext(read('method-cards.js'),ctx);const T=wpsTools(),ids=new Set(T.list('wps').map(x=>x.id));
  const runs=ctx.MethodCards.cards.filter(c=>c.run);assert.ok(runs.length>=14);
  const et=new Set(T.list('et').map(x=>x.id));for(const c of runs){assert.ok(c.host==='wps'?ids.has(c.run):et.has(c.run),c.id+' → '+c.run);}
  assert.equal(ctx.MethodCards.cards.find(c=>c.id==='mcW18').run,'deleteBlankPages');
});

function catalog(favorites=[]){
  const T=wpsTools();const ctx={window:{},OfficeTools:T,WorkbenchSettings:{snapshot:()=>({settings:{toolkit:{favorites,bindings:[]}}})}};ctx.window=ctx;
  T.isHidden=()=>false;
  vm.runInNewContext(read('method-cards.js')+'\n'+read('quick-catalog.js')+'\n'+read('quick-ribbon.js'),ctx);ctx.OfficeTools.host=()=>'wps';return ctx;
}
test('技巧导航：有一键工具的卡片不再重复列出同名工具，文档整理单独分组',()=>{
  const ctx=catalog(),list=ctx.QuickCatalog.list('wps');
  assert.equal(list.filter(x=>x.id==='deleteBlankPages').length,0,'由 mcW18 代表');
  assert.equal(list.find(x=>x.id==='mcW18').run,'deleteBlankPages');
  const ja=list.find(x=>x.id==='businessFormatJa');assert.equal(ja.group,'文档整理');assert.equal(ja.kind,'existing');
});
test('功能区“技巧导航”菜单：收藏排在最前，一键工具直接执行，指引打开侧栏',()=>{
  const ctx=catalog(['mcW03','mcW18']);let ran=null,opened=null;ctx.OnAction=c=>{ran=c.Id;};ctx.OpenQuickNavigator=id=>{opened=id;};
  const menu=ctx.GetQuickMenu(),ids=[...menu.matchAll(/id="qh_([^"]+)"/g)].map(m=>m[1]);
  assert.deepEqual(ids.slice(0,2).sort(),['mcW03','mcW18']);assert.match(menu,/label="★ /);assert.ok(ids.includes('businessFormatJa'));
  ctx.OnQuickMenuAction({Id:'qh_mcW18'});assert.equal(ran,'deleteBlankPages');assert.equal(opened,null);
  ctx.OnQuickMenuAction({Id:'qh_mcW03'});assert.equal(opened,'mcW03');
  ctx.OnQuickMenuAction({Id:'qh_businessFormatJa'});assert.equal(ran,'businessFormatJa');
});
test('功能区“我的收藏”：最多 6 个，按收藏顺序，空位隐藏',()=>{
  const fav=['mcW18','businessFormatJa','mcW03','mcW11','mcW12','mcW20','mcW26'];const ctx=catalog(fav),F=ctx.QuickFavorites;
  const slots=F.favorites('wps');assert.equal(slots.length,6);assert.deepEqual([...slots.map(x=>x.id)],fav.slice(0,6));
  assert.equal(F.slot('fav1').action,'deleteBlankPages');assert.equal(F.slot('fav3').action,'');assert.equal(catalog([]).QuickFavorites.slot('fav1'),null);
});
test('功能区按钮：文档整理、收藏都有文字和图标；收藏位跟着收藏变化',()=>{
  const xml=read('ribbon.xml'),text=JSON.parse(/window\.OfficeRibbonText=(\{.*\});?/s.exec(read('ribbon-text.js'))[1]);
  for(const id of ['docGroup','docMore','favGroup','favHint','deleteBlankPages','officialFormatCn','businessFormatJa','tableBeautify','removeCjkSpace','fav1','fav6']){assert.ok(xml.includes('id="'+id+'"'),id);assert.ok(text[id]&&text[id].label,id);}
  for(const name of ['blankpage','official','jadoc','landscape','picture','bold','options','nobreak','watermark','pagejump','spacing','star'])assert.ok(fs.existsSync(new URL('../public/icons/'+name+'.svg',import.meta.url)),name);
  const main=read('main.js');assert.match(main,/id==='docGroup'\)return h==='wps'/);assert.match(main,/QuickFavorites\.slot\(id\)/);assert.match(main,/id==='favHint'/);
});

test('Office 快捷键：十个键位、默认工具都存在、清单只在够新的 Office 上加入共享运行时',()=>{
  const sc=JSON.parse(read('office/shortcuts.json'));assert.equal(sc.actions.length,10);assert.equal(sc.shortcuts.length,10);
  for(const s of sc.shortcuts){assert.match(s.key.default,/^Ctrl\+Alt\+Shift\+[0-9]$/);assert.ok(sc.actions.some(a=>a.id===s.action&&a.type==='ExecuteFunction'&&a.name));}
  const ctx={};ctx.window=ctx;vm.runInNewContext(read('office/office-docplus.js')+'\n'+read('sheet-logic.js')+'\n'+read('office/office-sheetplus.js')+'\n'+read('office/office-tools.js')+'\n'+read('office/office-keys.js'),ctx);
  for(const host of ['word','excel','powerpoint']){const ids=new Set(ctx.OfficeAddinTools.list(host).map(x=>x.id));ids.add('openPane');for(const [k,id] of Object.entries(ctx.OfficeKeys.DEFAULTS[host]))assert.ok(ids.has(id),host+' '+k+' '+id);}
  const assoc=[];ctx.OfficeKeys.associate({actions:{associate:(n,f)=>assoc.push(n)}},'word',()=>{});assert.deepEqual([...assoc],sc.actions.map(a=>a.id));
  const tpl=read('office/manifest-template.xml');for(const k of ['{{REQUIREMENTS}}','{{RUNTIMES}}','{{EXTENDED}}'])assert.ok(tpl.includes(k),k);assert.equal((tpl.match(/\{\{RUNTIMES\}\}\r?\n        <DesktopFormFactor>/g)||[]).length,3);
  const ps=fs.readFileSync(new URL('../scripts/package/Initialize.ps1',import.meta.url),'utf8');
  assert.match(ps,/shortcuts = \(\$build -ge 17928\)/);assert.match(ps,/\.Replace\('\{\{REQUIREMENTS\}\}', ''\)\.Replace\('\{\{RUNTIMES\}\}', ''\)\.Replace\('\{\{EXTENDED\}\}', ''\)/);assert.match(ps,/\[switch\]\$NoOfficeShortcuts/);
});
test('Office 快捷键：改键位后保存在侧栏设置里，恢复默认',async()=>{
  const ctx={};ctx.window=ctx;let state={settings:{toolkit:{}}};ctx.WorkbenchSettings={snapshot:()=>state,update:p=>{Object.keys(p).forEach(k=>{state.settings[k]=Object.assign({},state.settings[k],p[k]);});return Promise.resolve(state);}};
  vm.runInNewContext(read('office/office-keys.js'),ctx);const K=ctx.OfficeKeys;
  assert.equal(K.slots('word')['1'],'deleteBlankPages');await K.set('word','1','pictureFix');assert.equal(K.slots('word')['1'],'pictureFix');assert.equal(K.slots('excel')['1'],'spotlight');
  await K.set('word','2','');assert.equal(K.slots('word')['2'],'');await K.reset('word');assert.equal(K.slots('word')['1'],'deleteBlankPages');
  await assert.rejects(K.set('word','x','toc'),/编号/);
  assert.equal(K.choices('powerpoint',[{id:'makeSlides',form:'outline'},{id:'formatSlide',label:'统一'}]).some(x=>x.id==='makeSlides'),false);
});
