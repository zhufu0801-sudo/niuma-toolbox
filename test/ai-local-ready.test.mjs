import test from 'node:test';
import assert from 'node:assert/strict';
import {createAiClient} from '../lib/ai.mjs';
const store=c=>({get:async()=>c});
test('本机 AI 没开：3 秒内直接说明，不发正式请求；开着时正常回答',async()=>{
  const calls=[];const down=async(u)=>{calls.push(u);throw Object.assign(Error('connect ECONNREFUSED'),{name:'TypeError'});};
  const ai=createAiClient({store:store({provider:'ollama',baseUrl:'http://127.0.0.1:11434/v1',model:'qwen2.5:7b'}),fetchImpl:down});
  await assert.rejects(ai.chat([{role:'user',content:'你好'}]),e=>e.status===503&&/本机 AI（本机 Ollama）没有运行/.test(e.message));
  assert.deepEqual(calls,['http://127.0.0.1:11434/v1/models']);
  const up=async(u,i)=>new Response(JSON.stringify(u.endsWith('/models')?{data:[{id:'qwen2.5:7b'}]}:{choices:[{message:{content:'你好！'}}]}),{status:200});
  const ok=createAiClient({store:store({provider:'ollama',baseUrl:'http://127.0.0.1:11434/v1',model:'qwen2.5:7b'}),fetchImpl:up});
  assert.match(JSON.stringify(await ok.chat([{role:'user',content:'你好'}])),/你好！/);
});
test('“不使用 AI”：清空服务商后面板显示未接入，Key 保留',async t=>{
  const fs=await import('node:fs/promises'),os=await import('node:os'),path=await import('node:path');const {createAiStore}=await import('../lib/ai.mjs');
  const d=await fs.mkdtemp(path.join(os.tmpdir(),'ai-off-'));t.after(()=>fs.rm(d,{recursive:true,force:true}));const s=createAiStore(path.join(d,'ai.json'));
  await s.save({provider:'ollama',model:'qwen2.5:7b'});assert.equal((await s.publicView()).config.ready,true);
  const v=await s.save({provider:'none'});assert.equal(v.config.provider,'');assert.equal(v.config.ready,false);
});
