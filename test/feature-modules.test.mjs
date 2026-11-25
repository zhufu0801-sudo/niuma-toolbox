// 代码模块化 + 功能开关：登记表、按宿主列出、隐藏后侧栏和功能区同步。
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
function tools(store={}){const storage={getItem:k=>store[k]??null,setItem:(k,v)=>{store[k]=v;}};const window={Application:{PluginStorage:storage},confirm:()=>true,localStorage:storage};vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date});return {T:window.OfficeTools,store};}

test('每个宿主列出自己的功能，表格里的“匹配周围格式”不进常用列表',()=>{
  const {T}=tools();
  const ids=h=>T.list(h).filter(x=>x.panel).map(x=>x.id);
  assert.ok(ids('wps').includes('heading1')&&ids('wps').includes('matchFormat'));
  assert.ok(ids('et').includes('formatHeader')&&!ids('et').includes('matchFormat'));
  assert.ok(ids('wpp').includes('formatSlide')&&!ids('wpp').includes('heading1'));
  assert.equal(T.describe('matchFormat','wpp').label,'匹配本页格式');
  assert.equal(T.describe('matchFormat','wps').label,'匹配周围格式');
});

test('同一功能 id 在不同模块登记不同宿主，会合并',()=>{
  const {T}=tools();
  for(const h of ['wps','et','wpp'])assert.ok(T.describe('insertTransfer',h).available,h);
  assert.ok(T.describe('replaceText','wps').available&&T.describe('replaceText','et').available&&!T.describe('replaceText','wpp').available);
});

test('隐藏功能会保存，并且跨宿主共享',()=>{
  const store={};const a=tools(store).T;a.setHidden(['heading1','formatHeader']);
  const b=tools(store).T;assert.equal(JSON.stringify(b.hidden()),JSON.stringify(['heading1','formatHeader']));assert.ok(b.isHidden('heading1'));assert.ok(b.prefsStamp()>0);
});

test('功能区：隐藏的按钮不显示，开关变化后刷新功能区',async()=>{
  const main=fs.readFileSync(new URL('../public/main.js',import.meta.url),'utf8');
  let hiddenIds=[],stamp=1,invalidated=0,tick;const window={};
  vm.runInNewContext(main,{window,setInterval:fn=>{tick=fn;},Date,Promise,alert:()=>{},location:{origin:'o',pathname:'/addon/wps/'},
    OfficeTools:{host:()=>'wps',isHidden:id=>hiddenIds.includes(id),prefsStamp:()=>stamp},ToolIcons:{forAction:()=>''},PaneBridge:{view:v=>v},StudioLaunch:{}});
  window.OnAddinLoad({Invalidate(){invalidated++;}});
  assert.equal(window.GetVisible({Id:'heading1'}),true);
  hiddenIds=['heading1'];stamp=2;tick();
  assert.equal(invalidated,1);assert.equal(window.GetVisible({Id:'heading1'}),false);
  assert.equal(window.GetVisible({Id:'wordGroup'}),true);assert.equal(window.GetVisible({Id:'sheetGroup'}),false);
  tick();assert.equal(invalidated,1,'没有变化时不重复刷新');
});
