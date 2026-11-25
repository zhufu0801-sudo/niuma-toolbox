import test from 'node:test';
import assert from 'node:assert/strict';
import {TOOL_MODULES} from './helpers/tools-code.mjs';
import {createClock,createDocument,createApp,createSharedBrowser,createFetch,loadScripts,flush} from './helpers/frontend-fakes.mjs';

const ORIGIN='http://127.0.0.1:39127';
const icons={path:id=>'icons/'+id+'.svg',forAction:id=>'icons/'+id+'.svg'};

// ---------- 功能区入口与侧栏页面一致 ----------
function openWorld({broadcast=true,localStorage=true,app=createApp(),host='wpp'}={}){
  const clock=createClock(),shared=createSharedBrowser({broadcast,localStorage}),alerts=[];
  const officeTools={host:()=>host,app:()=>app,run:async()=>'完成'};
  const ribbon=loadScripts({...shared,...clock,alert:m=>alerts.push(m),location:{origin:ORIGIN,pathname:`/addon/${host}/`,search:''},ToolIcons:icons,OfficeTools:officeTools,open:()=>null,addEventListener(){},removeEventListener(){}},['pane-bridge.js','studio-launch.js','main.js']);
  // 侧栏：真实的 tools-ui.js 运行在假 DOM 上。
  const doc=createDocument(),views=['quick','create','files'].map(view=>{const b=doc.createElement('button');b.dataset.view=view;return b;});
  const panels=['quick','create','create','files'].map((view,i)=>{const p=doc.createElement('section');p.dataset.panel=view;if(i===2)p.id='creation';return p;});
  doc.group('[data-view]',views);doc.group('[data-panel]',panels);doc.group('[data-action]',[]);doc.group('[data-jump]',[]);doc.group('button,textarea,input',[]);
  const replaced=[];
  const pane={...shared,...clock,document:doc,history:{replaceState:(a,b,url)=>replaced.push(url)},location:{origin:ORIGIN,pathname:'/tools.html',search:'?host='+host+'&view=quick'},URLSearchParams,ToolIcons:icons,Application:app,open:()=>null,addEventListener(){},removeEventListener(){}};
  return {ribbon,pane,doc,panels,app,alerts,clock,replaced,mountPane(){loadScripts(pane,[...TOOL_MODULES,'pane-bridge.js','studio-launch.js','tools-ui.js']);}};
}
const visiblePanel=w=>[...new Set(w.panels.filter(p=>!p.hidden).map(p=>p.dataset.panel))];

test('ribbon: first click creates the pane on the requested view',async()=>{
  const w=openWorld();
  w.ribbon.OnAction({Id:'makeSlides'});await flush();
  assert.equal(w.app.panes.length,1);assert.match(w.app.panes[0].url,/tools\.html\?host=wpp&view=create$/);assert.equal(w.app.panes[0].Visible,true);
  assert.deepEqual(w.alerts,[]);
});
for(const [name,options] of [['BroadcastChannel',{}],['PluginStorage only (no broadcast, no localStorage)',{broadcast:false,localStorage:false}]])
test(`ribbon: clicking another entry switches the open pane in place and keeps typed outline (${name})`,async()=>{
  const w=openWorld(options);w.mountPane();
  w.ribbon.OnAction({Id:'toolsOpen'});await flush();assert.equal(w.app.panes.length,1);
  w.doc.getElementById('content').value='项目概览\n项目背景';
  assert.deepEqual(visiblePanel(w),['quick']);
  w.ribbon.OnAction({Id:'makeSlides'});await w.clock.advance(600);
  assert.deepEqual(visiblePanel(w),['create']);
  assert.equal(w.doc.getElementById('pageTitle').textContent,'大纲制作');
  assert.equal(w.doc.getElementById('content').value,'项目概览\n项目背景','outline typed by the user must survive');
  assert.equal(w.app.panes.length,1);assert.equal(w.app.panes[0].deleted,false);
  assert.match(w.replaced.at(-1),/view=create$/);
  w.ribbon.OnAction({Id:'toolsOpen'});await w.clock.advance(600);
  assert.deepEqual(visiblePanel(w),['quick']);assert.equal(w.app.panes.length,1);
  assert.deepEqual(w.alerts,[]);
});
test('ribbon: an unresponsive pane retains user input and explains the manual navigation',async()=>{
  const w=openWorld();w.ribbon.OnAction({Id:'toolsOpen'});await flush();
  w.doc.getElementById('content').value='保留我的大纲';
  w.ribbon.OnAction({Id:'makeSlides'});await w.clock.advance(1700);
  assert.equal(w.app.panes.length,1);assert.equal(w.app.panes[0].deleted,false);
  assert.equal(w.doc.getElementById('content').value,'保留我的大纲');
  assert.equal(w.alerts.length,1);assert.match(w.alerts[0],/已保留填写内容/);
});

test('ribbon: a closed or invalid pane object is replaced instead of silently ignored',async()=>{
  const w=openWorld();w.ribbon.OnAction({Id:'toolsOpen'});await flush();
  Object.defineProperty(w.app.panes[0],'Visible',{set(){throw Error('pane closed');},get(){return false;}});
  w.ribbon.OnAction({Id:'makeSlides'});await flush();
  assert.equal(w.app.panes.length,2);assert.match(w.app.panes[1].url,/view=create$/);
});
test('ribbon: if an unresponsive pane cannot be rebuilt the user is told what to click',async()=>{
  const w=openWorld();w.ribbon.OnAction({Id:'toolsOpen'});await flush();
  delete w.app.panes[0].Delete;
  w.ribbon.OnAction({Id:'makeSlides'});await w.clock.advance(1700);
  assert.equal(w.alerts.length,1);assert.match(w.alerts[0],/制作/);assert.equal(w.app.panes.length,1);
});

// ---------- 版式工作室入口：区分 WPS 与浏览器预览 ----------
test('studio launch: browser preview opens a tab, and a blocked popup is reported rather than ignored',()=>{
  const w=openWorld();w.mountPane();const l=w.pane.StudioLaunch;
  const opened=[];
  let r=l.open({app:null,origin:ORIGIN,open:(url,target)=>{opened.push([url,target]);return {};}});
  assert.equal(r.ok,true);assert.equal(r.mode,'browser');assert.deepEqual(opened,[[ORIGIN+'/studio.html','_blank']]);assert.match(r.message,/不会修改 WPS/);
  r=l.open({app:null,origin:ORIGIN,open:()=>null});
  assert.equal(r.ok,false);assert.match(r.message,/拦截/);assert.equal(r.url,ORIGIN+'/studio.html');
});
test('studio launch: a real WPS error is surfaced and is not masked by opening a browser tab',()=>{
  const w=openWorld();w.mountPane();const l=w.pane.StudioLaunch;let tabs=0;
  const app={ShowDialog(){throw Error('窗口创建失败 0x80004005');}};
  const r=l.open({app,origin:ORIGIN,open:()=>{tabs++;return {};}});
  assert.equal(r.ok,false);assert.equal(r.mode,'wps');assert.match(r.message,/0x80004005/);assert.equal(tabs,0);
  const old=l.open({app:{},origin:ORIGIN,open:()=>{tabs++;return {};}});
  assert.equal(old.ok,false);assert.equal(old.kind,'unsupported');assert.equal(tabs,0);
  const good=createApp();assert.equal(l.open({app:good,origin:ORIGIN,open:()=>null}).ok,true);assert.equal(good.dialogs[0][0],ORIGIN+'/studio.html');
});
test('side pane: launch button shows a visible failure notice and a manual link',()=>{
  const w=openWorld({app:{ShowDialog(){throw Error('boom');}}});w.mountPane();
  w.doc.getElementById('studioLaunch').click();
  const notice=w.doc.getElementById('studioNotice'),link=w.doc.getElementById('studioFallback');
  assert.equal(notice.hidden,false);assert.match(notice.textContent,/boom/);assert.ok(notice.classList.contains('error'));
  assert.equal(link.hidden,false);assert.equal(link.href,ORIGIN+'/studio.html');
  assert.equal(w.doc.getElementById('loadHint').removed,true,'load hint is removed once the page initialised');
});
test('ribbon studio button: failure is announced with a manual address',async()=>{
  const w=openWorld({app:{ShowDialog(){throw Error('boom');}}});
  w.ribbon.OnAction({Id:'studioOpen'});
  assert.equal(w.alerts.length,1);assert.match(w.alerts[0],/boom/);assert.match(w.alerts[0],/127\.0\.0\.1:39127\/studio\.html/);
});

// ---------- studio.js 异步流程 ----------
function studioWorld(){
  const clock=createClock(),doc=createDocument(),net=createFetch(),downloads=[];
  const fields=['deckTitle','subtitle','outline'].map(id=>doc.getElementById(id)),radio=doc.createElement('input');radio.type='radio';radio.value='business';radio.checked=true;
  const example=doc.getElementById('example');
  doc.group('input[name=theme]:checked',[radio]);doc.group('.theme',[]);doc.group('input,textarea',[...fields,radio]);doc.group('input,textarea,#example',[...fields,radio,example]);
  const win=loadScripts({document:doc,fetch:net.fetch,AbortController,Blob,URLSearchParams,location:{search:''},URL:{createObjectURL:b=>{downloads.push(b);return 'blob:x';},revokeObjectURL(){}},...clock,addEventListener(){}},['studio.js']);
  fields[0].value='季度汇报';fields[2].value='成果\n完成事项';
  const plan=(pages=1)=>({width:1280,height:720,warnings:[],pages:Array.from({length:pages},(_,i)=>({role:i?'content':'cover',background:'123749',elements:[{type:'rect',x:0,y:0,w:10,h:10,fill:'FFFFFF'},{type:'text',x:1,y:1,w:5,h:5,fontSize:20,lines:['文'],bold:true,color:'FFFFFF'}]}))});
  const el=id=>doc.getElementById(id),edit=(id,value)=>{el(id).value=value;el(id).dispatch('input');};
  const preview=async()=>{el('studioForm').onsubmit({preventDefault(){}});await flush();};
  return {win,doc,net,clock,el,edit,plan,preview,fields,example,downloads};
}
const unlocked=s=>s.fields.every(f=>!f.disabled)&&!s.el('previewButton').disabled;

test('studio: a preview of outdated content is discarded and export stays locked until the current content is previewed',async()=>{
  const s=studioWorld();
  assert.equal(s.el('loadHint').removed,true);
  await s.preview();const first=s.net.studio('preview')[0];
  s.edit('outline','成果\n完成事项\n新增一行');           // 预览返回前继续编辑
  first.respond(200,s.plan(1));await flush();
  assert.equal(s.el('download').disabled,true);assert.equal(s.el('openWps').disabled,true);assert.equal(s.el('pages').children.length,0,'outdated plan must not be rendered');
  await s.preview();s.net.studio('preview')[1].respond(200,s.plan(2));await flush();
  assert.equal(s.el('download').disabled,false);assert.equal(s.el('openWps').disabled,false);assert.equal(s.el('pages').children.length,2);
  s.edit('deckTitle','改名后的标题');                      // 已有预览后再改：旧预览不可导出，并标为待更新
  assert.equal(s.el('download').disabled,true);assert.equal(s.el('openWps').disabled,true);
  assert.ok(s.el('pages').classList.contains('stale'));assert.match(s.el('pageCount').textContent,/待更新/);
  await s.clock.advance(700);s.net.studio('preview')[2].respond(200,s.plan(3));await flush();
  assert.equal(s.el('download').disabled,false);assert.equal(s.el('pages').classList.contains('stale'),false);assert.equal(s.el('pageCount').textContent,'3 页');
});
test('studio: a hung preview times out with a clear message, controls recover and retry works',async()=>{
  const s=studioWorld();await s.preview();
  assert.equal(s.el('previewButton').disabled,true);assert.match(s.el('status').textContent,/正在整理/);
  await s.clock.advance(5000);assert.match(s.el('status').textContent,/较慢/);
  await s.clock.advance(15100);
  assert.ok(s.el('status').classList.contains('error'));assert.match(s.el('status').textContent,/用时过长.*重试/);assert.equal(s.el('previewButton').disabled,false);
  await s.preview();s.net.studio('preview')[1].respond(200,s.plan(1));await flush();
  assert.equal(s.el('download').disabled,false);
});
test('studio: a service that is down gives a recovery hint instead of a raw error',async()=>{
  const s=studioWorld();s.win.fetch=()=>Promise.reject(new TypeError('Failed to fetch'));
  await s.preview();assert.match(s.el('status').textContent,/无法连接本地服务/);assert.equal(s.el('previewButton').disabled,false);
});
async function readyStudio(){const s=studioWorld();await s.preview();s.net.studio('preview')[0].respond(200,s.plan(2));await flush();return s;}
test('studio: double clicking "open in WPS" sends one request, locks inputs meanwhile, and unlocks after a non-JSON server failure',async()=>{
  const s=await readyStudio();
  s.el('openWps').click();s.el('openWps').click();await flush();
  assert.equal(s.net.studio('open').length,1);assert.equal(s.fields.every(f=>f.disabled),true);assert.equal(s.el('download').disabled,true);assert.equal(s.el('openWps').disabled,true);
  s.net.studio('open')[0].respond(502,undefined);await flush();               // 返回体不是 JSON
  assert.ok(s.el('status').classList.contains('error'));assert.match(s.el('status').textContent,/处理失败，请重试/);
  assert.equal(unlocked(s),true);assert.equal(s.el('openWps').disabled,false,'same revision, so preview is still exportable');
  s.el('openWps').click();await flush();s.net.studio('open')[1].respond(200,{opened:true,pages:2,filePath:'C:/x/new.pptx'});await flush();
  assert.match(s.el('status').textContent,/已启动 WPS 演示.*2 页/);assert.equal(unlocked(s),true);
});
test('studio: waiting for WPS too long stops with advice not to blindly retry, and unlocks',async()=>{
  const s=await readyStudio();s.el('openWps').click();await flush();
  await s.clock.advance(90100);
  assert.match(s.el('status').textContent,/WPS 中是否已出现新演示/);assert.equal(unlocked(s),true);assert.ok(s.el('status').classList.contains('error'));
});
test('studio: an expired token is refreshed once, and server error text is shown as-is',async()=>{
  const s=await readyStudio();s.el('download').click();await flush();
  const first=s.net.studio('export')[0];assert.equal(first.options.headers['X-Coach-Token'],'token-1');first.respond(403,{error:'本机连接已过期，请重新尝试。'});await flush();
  const second=s.net.studio('export')[1];assert.equal(second.options.headers['X-Coach-Token'],'token-2');
  second.respond(200,{blob:true});await flush();
  assert.match(s.el('status').textContent,/PPTX 已生成/);assert.equal(s.downloads.length,1);
  const link=s.doc.created.find(e=>e.download);assert.equal(link.download,'季度汇报.pptx');assert.equal(unlocked(s),true);
});
