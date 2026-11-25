import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const read=f=>fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8');
function load(){const w={};const ctx={window:w,OfficeTools:{list:()=>[{id:'formatHeader',label:'美化表头',desc:'d',panel:true}],isHidden:()=>false,host:()=>'et'},WorkbenchSettings:{snapshot:()=>({settings:{toolkit:{bindings:[]}}})}};ctx.window=w;w.OfficeTools=ctx.OfficeTools;w.WorkbenchSettings=ctx.WorkbenchSettings;
 for(const f of ['method-cards.js','quick-catalog.js','quick-ribbon.js'])vm.runInNewContext(read(f),{...ctx,window:w,OfficeTools:ctx.OfficeTools,WorkbenchSettings:ctx.WorkbenchSettings,QuickCatalog:w.QuickCatalog});return w;}
test('办公方法卡：65 张，按应用分配，只做操作指引',()=>{
 const w=load(),cards=w.MethodCards.cards;
 assert.equal(cards.length,65);
 assert.equal(new Set(cards.map(c=>c.id)).size,65);
 for(const c of cards){assert.match(c.id,/^[A-Za-z][A-Za-z0-9]{0,79}$/);assert.equal(c.kind,'guide');assert.ok(c.label&&c.desc&&c.steps&&c.note,c.id);assert.match(c.note,/^来源：/);if(c.source)assert.match(c.source,/^https:\/\/www\.douyin\.com\/video\/\d+$/);}
 assert.equal(w.MethodCards.forHost('et').length,36);assert.equal(w.MethodCards.forHost('wps').length,36);assert.equal(w.MethodCards.forHost('wpp').length,7);
});
test('技巧导航包含方法卡，功能区下拉菜单不列出方法卡',()=>{
 const w=load();const list=w.QuickCatalog.list('et');
 assert.ok(list.some(x=>x.id==='mcE04'));assert.ok(list.some(x=>x.id==='selectionStats'));
 const menu=w.GetQuickMenu();assert.ok(!menu.includes('qh_mcE04'));assert.ok(menu.includes('qh_selectionStats'));
});
test('日文补充词典覆盖全部方法卡、技巧导航和新工具',()=>{
 const r={};r.window=r;vm.runInNewContext(read('i18n-ja.js'),{window:r});vm.runInNewContext(read('i18n-ja-extra.js'),{window:r});
 const J=r.OfficeI18nJa,w=load();
 for(const c of w.MethodCards.cards)for(const k of ['label','desc','steps','note'])assert.ok(Object.hasOwn(J,c[k]),c.id+' '+k);
 for(const x of w.QuickCatalog.items)assert.ok(Object.hasOwn(J,x.label),x.label);
 for(const k of ['跨列居中','缩小字体填充','空白填上一格','图片适应单元格','图片统一大小','图片不超出页面','证件照按实际尺寸','中文公式助手','VBA 宏助手'])assert.ok(Object.hasOwn(J,k),k);
 for(const f of ['tools.html','index.html','settings.html','start.html'])assert.match(read(f),/i18n-ja\.js"><\/script><script src="i18n-ja-extra\.js">/);
});
