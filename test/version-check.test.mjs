// WPS 不重启时会保留旧脚本：按钮必须提示重启，而不是静默执行旧代码。
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const main=fs.readFileSync(new URL('../public/main.js',import.meta.url),'utf8');
const flush=()=>new Promise(r=>setTimeout(r,0));
function load(serverVersion){
  const alerts=[],runs=[];const window={};
  const ctx={window,alert:m=>alerts.push(m),location:{origin:'http://127.0.0.1:39127',pathname:'/addon/wps/'},Date,Promise,
    fetch:async url=>{assert.equal(url,'version.json');return {ok:true,json:async()=>({version:serverVersion})};},
    OfficeTools:{version:'0.3.3',host:()=>'wps',app:()=>({}),runFromUi:async id=>{runs.push(id);return '完成';}},
    ToolIcons:{forAction:()=>''},PaneBridge:{view:v=>v},StudioLaunch:{}};
  vm.runInNewContext(main,ctx);return {ribbon:window,alerts,runs};
}
test('版本一致时正常执行',async()=>{const w=load('0.3.3');w.ribbon.OnAction({Id:'matchFormat'});await flush();await flush();assert.deepEqual(w.runs,['matchFormat']);assert.equal(w.alerts.length,0);});
test('WPS 里是旧脚本时提示重启，且不执行，每次点击都会提示',async()=>{
  const w=load('9.9.9');
  w.ribbon.OnAction({Id:'matchFormat'});await flush();await flush();
  w.ribbon.OnAction({Id:'heading1'});await flush();await flush();
  assert.deepEqual(w.runs,[]);assert.equal(w.alerts.length,2);assert.match(w.alerts[0],/重新打开/);
});
test('功能区执行后没有需要修改的内容时弹出说明，避免用户以为按钮坏了',async()=>{
  const w=load('0.3.3');w.ribbon.OnAction({Id:'x'});await flush();await flush();
  const alerts=[];const ctx2={};
  const main2=fs.readFileSync(new URL('../public/main.js',import.meta.url),'utf8');
  const win={};vm.runInNewContext(main2,{window:win,alert:m=>alerts.push(m),location:{origin:'o',pathname:'/addon/wps/'},Date,Promise,fetch:async()=>({ok:true,json:async()=>({version:'0.3.3'})}),OfficeTools:{version:'0.3.3',host:()=>'wps',app:()=>({}),runFromUi:async()=>'选区内没有可删除的完整空段落。'},ToolIcons:{forAction:()=>''},PaneBridge:{view:v=>v},StudioLaunch:{}});
  win.OnAction({Id:'removeBlankParas'});await flush();await flush();await flush();
  assert.deepEqual(alerts,['选区内没有可删除的完整空段落。']);
});
