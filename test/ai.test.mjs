import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {PROVIDERS,checkBaseUrl,createAiStore,createAiClient,parseOutline,outlinePrompt,maskKey} from '../lib/ai.mjs';
import {planStudio} from '../lib/studio.mjs';
import {createCoachServer} from '../server.mjs';
const tmp=async t=>{const d=await fs.mkdtemp(path.join(os.tmpdir(),'ai-test-'));t.after(()=>fs.rm(d,{recursive:true,force:true}));return d;};
test('AI 服务商：每家都有官网申请页，地址必须是 https（本机 Ollama 除外）',()=>{
  for(const p of PROVIDERS.filter(x=>x.id!=='custom')){assert.match(p.keyUrl,/^https:\/\//,p.id);if(p.baseUrl)assert.equal(checkBaseUrl(p.baseUrl),p.baseUrl.replace(/\/+$/,''));}
  assert.throws(()=>checkBaseUrl('http://api.example.com/v1'),/https/);assert.throws(()=>checkBaseUrl('https://u:p@api.example.com'),/账号/);assert.throws(()=>checkBaseUrl('https://api.example.com/v1?key=1'));
  assert.equal(checkBaseUrl('http://127.0.0.1:11434/v1/'),'http://127.0.0.1:11434/v1');assert.equal(maskKey('sk-1234567890abcd'),'sk-…abcd');
});
test('AI 设置：密钥只存本机文件，对外只给打码提示；不重填时保留旧密钥，可清除',async t=>{
  const d=await tmp(t),store=createAiStore(path.join(d,'ai.json'));
  const v=await store.save({provider:'deepseek',model:'deepseek-flash',apiKey:'sk-secret-123456'});
  assert.equal(JSON.stringify(v).includes('sk-secret'),false);assert.equal(v.config.keyHint,'sk-…3456');assert.equal(v.config.ready,true);assert.equal(v.config.baseUrl,'https://api.deepseek.com');
  await store.save({provider:'deepseek',model:'deepseek-v4-pro'});assert.equal((await store.get()).apiKey,'sk-secret-123456');
  await assert.rejects(store.save({provider:'deepseek',model:'x',apiKey:'sk with space'}),/空格/);
  await store.save({provider:'deepseek',model:'x',clearKey:true});assert.equal((await store.get()).apiKey,'');
  await assert.rejects(store.save({provider:'nope'}),/服务商/);
});
test('AI 调用：Bearer 密钥、OpenAI 兼容路径、不带 temperature；错误给出能看懂的原因并在模型不对时列出可选模型',async t=>{
  const d=await tmp(t),store=createAiStore(path.join(d,'ai.json'));await store.save({provider:'claude',model:'claude-sonnet-5-5',apiKey:'sk-ant-abcdef'});
  const calls=[];let mode='ok';
  const fetchImpl=async(url,init)=>{calls.push({url,init});
    if(url.endsWith('/models'))return new Response(JSON.stringify({data:[{id:'claude-haiku-5-5'},{id:'claude-sonnet-5-5'}]}),{status:200});
    if(mode==='401')return new Response('{"error":{"message":"invalid x-api-key"}}',{status:401});
    if(mode==='404')return new Response('{"error":{"message":"model not found"}}',{status:404});
    return new Response(JSON.stringify({choices:[{message:{content:'<think>…</think>OK'}}]}),{status:200});};
  const ai=createAiClient({store,fetchImpl});
  const r=await ai.chat([{role:'user',content:'hi'}]);assert.equal(r.answer,'OK');
  assert.equal(calls[0].url,'https://api.anthropic.com/v1/chat/completions');assert.equal(calls[0].init.headers.Authorization,'Bearer sk-ant-abcdef');
  const body=JSON.parse(calls[0].init.body);assert.equal(body.model,'claude-sonnet-5-5');assert.equal('temperature' in body,false);
  mode='401';await assert.rejects(ai.chat([{role:'user',content:'x'}]),/API Key 不正确/);
  mode='404';const tr=await ai.test();assert.equal(tr.ok,false);assert.deepEqual(tr.models,['claude-haiku-5-5','claude-sonnet-5-5']);
  const none=createAiClient({store:createAiStore(path.join(d,'none.json')),fetchImpl});await assert.rejects(none.chat([]),/还没有设置 AI/);
  const off=createAiClient({store,fetchImpl:async()=>{throw new TypeError('fetch failed');}});mode='ok';await assert.rejects(off.chat([{role:'user',content:'x'}]),/连不上 AI 服务/);
});
test('AI 大纲：从回答里取出 JSON，转成带版式标记的大纲，模板能直接排版（中日文）',()=>{
  const answer='好的：\n```json\n{"title":"新员工培训","subtitle":"第一周","sections":[{"heading":"培训安排","layout":"timeline","points":["第1周：熟悉制度","第2周：岗位实操","第3周：考核"]},{"heading":"核心数据","layout":"stat","points":["95% 通过率","30 门课程"]},{"heading":"四个原则","layout":"grid","points":["安全：先安全后效率","质量：一次做对","协作：主动沟通","成长：每周复盘"]},{"heading":"寄语","layout":"quote","points":["把每一件小事做好。"]}]}\n```';
  const o=parseOutline(answer,'zh');assert.equal(o.title,'新员工培训');assert.match(o.outline,/培训安排 \[时间线\]/);assert.match(o.outline,/核心数据 \[数据\]/);
  const plan=planStudio({...o,theme:'coral'});assert.deepEqual(plan.pages.filter(p=>p.role==='content').map(p=>p.layout),['timeline','stat','grid','quote']);
  const ja=parseOutline('{"title":"新製品発表","sections":[{"heading":"特長","layout":"grid","points":["速い：処理が2倍","軽い：重量30%減","静か：騒音を半分に"]}]}','ja');assert.match(ja.outline,/特長 \[カード\]/);
  const jp=planStudio({...ja,theme:'wa'});assert.equal(jp.lang,'ja');assert.ok(jp.pages.some(p=>p.elements.some(e=>e.face==='Yu Mincho')));assert.ok(jp.pages.at(-1).elements.some(e=>e.text==='ご清聴ありがとうございました'));
  assert.throws(()=>parseOutline('抱歉，我不能','zh'),/格式/);assert.throws(()=>parseOutline('{"sections":[]}','zh'),/空/);
  assert.match(outlinePrompt({topic:'x',lang:'ja'})[1].content,/日本語/);
});
test('AI 接口：读取设置不含密钥，写入需要本机口令，生成大纲走用户选的服务',async t=>{
  const d=await tmp(t);let sent=null;
  const aiFetch=async(url,init)=>{sent={url,init};return new Response(JSON.stringify({choices:[{message:{content:'{"title":"季度汇报","sections":[{"heading":"成果","points":["完成 A","完成 B"]}]}'}}]}),{status:200});};
  const server=createCoachServer({aiFile:path.join(d,'ai.json'),aiFetch,cardsFile:path.join(d,'c.json'),settingsFile:path.join(d,'s.json'),backupDir:d});await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await (await fetch(base+'/api/session')).json()).token;
  const post=(p,b,k=token)=>fetch(base+p,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':k},body:JSON.stringify(b)});
  assert.equal((await post('/api/ai/config',{provider:'qwen',model:'qwen-plus',apiKey:'sk-qwen-999999'},'bad')).status,403);
  assert.equal((await post('/api/ai/ppt-outline',{topic:'季度汇报'})).status,409);
  assert.equal((await post('/api/ai/config',{provider:'qwen',model:'qwen-plus',apiKey:'sk-qwen-999999'})).status,200);
  const view=await (await fetch(base+'/api/ai/config')).text();assert.ok(!view.includes('sk-qwen-999999'));assert.ok(view.includes('sk-…9999'));
  const r=await (await post('/api/ai/ppt-outline',{topic:'季度汇报',pages:4})).json();assert.equal(r.title,'季度汇报');assert.match(r.outline,/成果\n完成 A/);
  assert.equal(sent.url,'https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions');assert.equal(sent.init.headers.Authorization,'Bearer sk-qwen-999999');
  assert.equal((await post('/api/ai/ppt-outline',{topic:''})).status,400);
});
test('思考型模型：正文为空时测试连接仍算连通，生成时给出明确原因；本机模型等待更久',async t=>{
  const d=await fs.mkdtemp(path.join(os.tmpdir(),'ai-think-'));t.after(()=>fs.rm(d,{recursive:true,force:true}));
  const store=createAiStore(path.join(d,'ai.json'));await store.save({provider:'ollama',model:'deepseek-r1:8b'});let seen;
  const ai=createAiClient({store,fetchImpl:async(url,init)=>{if(url.endsWith('/models'))return new Response('{"data":[]}',{status:200});seen=JSON.parse(init.body);return new Response(JSON.stringify({choices:[{finish_reason:'length',message:{content:'',reasoning:'嗯，用户让我…'}}]}),{status:200});}});
  const r=await ai.test();assert.equal(r.ok,true);assert.equal(r.thinking,true);assert.equal(seen.max_tokens,512);
  await assert.rejects(ai.chat([{role:'user',content:'x'}]),/思考/);
  const ok=createAiClient({store,fetchImpl:async()=>new Response(JSON.stringify({choices:[{message:{content:'<think>先想想</think>\n{"a":1}'}}]}),{status:200})});assert.equal((await ok.chat([])).answer,'{"a":1}');
});
test('不使用 AI：可同时清除 Key、修改个人信息保护方式；不带 clearKey 时 Key 保留',async t=>{
  const d=await tmp(t),store=createAiStore(path.join(d,'ai.json'));
  await store.save({provider:'deepseek',model:'deepseek-flash',apiKey:'sk-secret-123456'});
  let v=await store.save({provider:'none',privacy:{mode:'remind'}});assert.equal(v.config.provider,'');assert.equal(v.config.hasKey,true);assert.equal(v.config.privacy.mode,'remind');assert.equal(v.config.ready,false);
  v=await store.save({provider:'none',clearKey:true});assert.equal(v.config.hasKey,false);assert.equal(v.config.privacy.mode,'remind');
});
