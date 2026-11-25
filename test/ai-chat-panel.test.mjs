import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const src=fs.readFileSync(new URL('../public/ai-chat-panel.js',import.meta.url),'utf8');
function load({read,answer,fail}={}){
  const stub=()=>({value:'',textContent:'',hidden:false,disabled:false,checked:false,className:'',children:[],append(...x){this.children.push(...x);},appendChild(x){this.children.push(x);},setAttribute(k,v){this[k]=v;},addEventListener(){},onclick:null});
  const els={};const section={className:'',dataset:{},set innerHTML(h){for(const m of h.matchAll(/id="(ac[A-Za-z]+)"/g))els[m[1]]=stub();},querySelector:s=>els[s.slice(1)]};
  const sent=[];const P={host:()=>({read:read||(async()=>({kind:'text',text:'选中的段落'})),insert:async()=>'已插入'}),api:async(url,body)=>{sent.push({url,body});if(fail)throw Error(fail);return {text:answer||'回答',model:'m1',privacy:{masked:0}};}};
  const w={OfficeTools:{lib:{ask:async()=>true}},AiTextPanel:P,fetch:async()=>({json:async()=>({config:{}})}),navigator:{},
    document:{createElement:t=>t==='section'?section:stub(),getElementById:id=>id==='aiTextPanel'?{parentNode:{insertBefore(){}}}:null,addEventListener(){},body:stub()}};
  w.window=w;vm.runInNewContext(src,w);return {w,els,sent};
}
test('AI 对话：发送问题、带上历史、可附带选中的内容；失败时把问题放回输入框',async()=>{
  const {w,els,sent}=load({answer:'用 VLOOKUP'});
  els.acInput.value='怎么查价格？';await w.AiChatPanel.send();
  assert.equal(sent[0].url,'/api/ai/chat');assert.deepEqual(JSON.parse(JSON.stringify(sent[0].body.messages)),[{role:'user',content:'怎么查价格？'}]);assert.equal(sent[0].body.context,'');
  assert.deepEqual([...w.AiChatPanel.history().map(m=>m.role)],['user','assistant']);assert.equal(els.acInput.value,'');assert.match(els.acStatus.textContent,/已回答（m1）/);
  els.acInput.value='再详细一点';els.acContext.checked=true;await w.AiChatPanel.send();
  assert.equal(sent[1].body.messages.length,3);assert.equal(sent[1].body.context,'选中的段落');assert.equal(w.AiChatPanel.history()[2].withContext,true);
  const f=load({fail:'AI 服务连接失败'});f.els.acInput.value='问题';await f.w.AiChatPanel.send();
  assert.equal(f.w.AiChatPanel.history().length,0);assert.equal(f.els.acInput.value,'问题');assert.match(f.els.acStatus.className,/error/);
  const e=load();await e.w.AiChatPanel.send();assert.equal(e.sent.length,0);assert.match(e.els.acStatus.textContent,/请先输入问题/);
  const c=load({read:async()=>({kind:'cells',address:'A1:B2',items:[{text:'编号'},{text:'价格'}]})});c.els.acInput.value='解释';c.els.acContext.checked=true;await c.w.AiChatPanel.send();assert.equal(c.sent[0].body.context,'A1:B2\n编号\n价格');
});
