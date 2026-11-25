import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs/promises';import os from 'node:os';import path from 'node:path';import vm from 'node:vm';
import {createMaskVault,maskEngine} from '../lib/masking.mjs';import {createMcpBridge} from '../lib/mcp-bridge.mjs';
const read=f=>fs.readFile(new URL('../public/'+f,import.meta.url),'utf8');
test('clipboard mappings encrypted, prune only oldest clips and newest code wins',async t=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'clip-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));const file=path.join(dir,'vault.json');
 const secrets={protect:async value=>({value,protection:'dpapi'}),unprotect:async value=>value};
 const vault=createMaskVault(file,{secrets});await vault.addRecord({kind:'doc',name:'keep'});
 for(let i=0;i<22;i++)await vault.addRecord({kind:'clip',name:'复制给 AI',mode:'code',pairs:[{original:'张三'+i,masked:'【名称A】'}]});
 const state=await vault.state();assert.equal(state.records.filter(x=>x.kind==='clip').length,20);assert.equal(state.records.filter(x=>x.kind==='doc').length,1);
 assert.ok(!(await fs.readFile(file,'utf8')).includes('张三'));const pairs=await vault.clipPairs();assert.equal(pairs[0].original,'张三21');assert.equal(pairs.at(-1).original,'张三2');
 assert.equal((await maskEngine()).restoreCodes('[名称A]',pairs).text,'张三21');
});
test('clipboard API, legacy copy, manual fallback and manual restore through actual handlers',async()=>{
 const nodes={},body={appendChild(){}},document={body,execCommand:()=>false,createElement:tag=>el(tag)};let task,inserted='',copied='';
 function el(tag,cls,text){return {tag,children:[],textContent:text||'',value:'',style:{},hidden:false,set id(v){this._id=v;nodes[v]=this;},get id(){return this._id;},append(...a){this.children.push(...a);},setAttribute(){},focus(){},select(){},remove(){}};}
 const c={document,setTimeout,clearTimeout,navigator:{clipboard:{readText:async()=>{throw Error('denied');},writeText:async s=>{copied=s;}}}};c.window=c;
 c.OfficeTools={host:()=> 'wps',app:()=>({Selection:{Text:'张三 13812345678',TypeText:s=>{inserted=s;}}})};
 vm.runInNewContext(await read('mask-engine.js')+'\n'+await read('mask-clipboard.js'),c);
 const vault=createMaskVault(null),panel=el('section');let notice='';
 const o={M:(z,j)=>z,el,say:s=>{notice=s;},task:fn=>{task=Promise.resolve().then(fn);},opts:()=>({words:c.MaskEngine.parseWords('张三').words,types:['PHONE']}),lang:()=> 'zh',refresh:async()=>{},post:async input=>input.action==='clip-pairs'?{pairs:await vault.clipPairs()}:{record:await vault.addRecord(input.record)}};
 c.MaskClipboard.mount(panel,o);nodes.mkClipCopy.onclick();await task;assert.equal(copied,'【名称A】 【电话1】');assert.ok(!copied.includes('张三'));
 nodes.mkClipPaste.onclick();await task;assert.match(notice,/读不到剪贴板/);assert.equal(inserted,'');
 nodes.mkClipInput.value='[ 名称A ] ［电话1］ 【名称Z】';panel.children[0].children.find(x=>x.textContent==='还原并插入').onclick();await task;
 assert.equal(inserted,'张三 13812345678 【名称Z】');assert.match(notice,/未知代号/);
 c.navigator.clipboard.writeText=async()=>{throw Error('denied');};let legacy=0;document.execCommand=()=>{legacy++;return true;};assert.equal(await c.MaskClipboard.copyText('safe'),true);assert.equal(legacy,1);
 document.execCommand=()=>false;assert.equal(await c.MaskClipboard.copyText('safe'),false);nodes.mkClipCopy.onclick();await task;assert.match(notice,/Ctrl\+C/);assert.equal(nodes.mkClipOutput.value,'【名称A】 【电话1】');
 c.navigator.clipboard.readText=()=>new Promise(()=>{});assert.equal(await c.MaskClipboard.readText(c,5),null);
});
test('MCP result refreshes heartbeat after a long native confirmation',async()=>{
 let now=0;const b=createMcpBridge({now:()=>now}),client={clientId:'0123456789abcdef',host:'wps',suite:'wps'};b.register(client);
 const pending=b.execute({...client,action:'insertText',args:{text:'qa'}}),task=b.poll(client).task;now=10000;b.finish({...client,id:task.id,result:'done'});await pending;
 const second=b.execute({...client,action:'readDocument'}),next=b.poll(client).task;b.finish({...client,id:next.id,result:'qa'});assert.equal(await second,'qa');b.close();
});
