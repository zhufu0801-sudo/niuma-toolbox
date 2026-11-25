// AI 接入：客户填自己的 API Key，工具箱在本机直接调用服务商的“OpenAI 兼容”接口。
// 密钥只保存在本机 generated\ai.json（Windows 上加密），不发给页面、不写进诊断信息；只连接用户选定的地址。
// 所有发出去的内容先经过隐私保护关卡（lib/privacy.mjs），收到的回答也经过它。
// 各家地址和申请页核对于 2026-11-11（以官网为准，模型名会更新，可点“获取模型列表”）。
import fs from 'node:fs/promises';
import path from 'node:path';
import { createSecretBox } from './secret.mjs';
import { PRIVACY_MODES } from './privacy.mjs';

export const PROVIDERS=[
  {id:'deepseek',name:'DeepSeek（深度求索）',nameJa:'DeepSeek',region:'cn',baseUrl:'https://api.deepseek.com',model:'deepseek-flash',keyUrl:'https://platform.deepseek.com/api_keys',docUrl:'https://api-docs.deepseek.com/zh-cn/'},
  {id:'qwen',name:'通义千问（阿里云百炼）',nameJa:'Qwen（Alibaba Cloud）',region:'cn',baseUrl:'https://dashscope.aliyuncs.com/compatible-mode/v1',model:'qwen-plus',keyUrl:'https://bailian.console.aliyun.com/model/settings/api-key',docUrl:'https://help.aliyun.com/zh/model-studio/get-api-key'},
  {id:'zhipu',name:'智谱 GLM',nameJa:'Zhipu GLM',region:'cn',baseUrl:'https://open.bigmodel.cn/api/paas/v4',model:'glm-5.3',keyUrl:'https://bigmodel.cn/usercenter/proj-mgmt/apikeys',docUrl:'https://docs.bigmodel.cn/cn/guide/start/quick-start'},
  {id:'kimi',name:'Kimi（月之暗面）',nameJa:'Kimi（Moonshot）',region:'cn',baseUrl:'https://api.moonshot.cn/v1',model:'',keyUrl:'https://platform.kimi.com/',docUrl:'https://platform.kimi.com/docs'},
  {id:'doubao',name:'豆包（火山方舟）',nameJa:'Doubao（Volcengine）',region:'cn',baseUrl:'https://ark.cn-beijing.volces.com/api/v3',model:'',keyUrl:'https://console.volcengine.com/ark',docUrl:'https://www.volcengine.com/docs/82379/1330626'},
  {id:'openai',name:'OpenAI（ChatGPT）',nameJa:'OpenAI（ChatGPT）',region:'global',baseUrl:'https://api.openai.com/v1',model:'',keyUrl:'https://platform.openai.com/api-keys',docUrl:'https://developers.openai.com/api/docs'},
  {id:'claude',name:'Anthropic Claude',nameJa:'Anthropic Claude',region:'global',baseUrl:'https://api.anthropic.com/v1',model:'claude-sonnet-5-5',keyUrl:'https://platform.claude.com/settings/keys',docUrl:'https://platform.claude.com/docs/zh-CN/cli-sdks-libraries/libraries/openai-sdk'},
  {id:'gemini',name:'Google Gemini',nameJa:'Google Gemini',region:'global',baseUrl:'https://generativelanguage.googleapis.com/v1beta/openai',model:'gemini-3.8-flash',keyUrl:'https://aistudio.google.com/apikey',docUrl:'https://ai.google.dev/gemini-api/docs/openai'},
  {id:'ollama',name:'本机 Ollama（免费、离线）',nameJa:'ローカル Ollama（無料・オフライン）',region:'local',baseUrl:'http://127.0.0.1:11434/v1',model:'',keyUrl:'https://ollama.com/download',docUrl:'https://ollama.com/',noKey:true},
  {id:'custom',name:'其他（OpenAI 兼容接口）',nameJa:'その他（OpenAI 互換）',region:'any',baseUrl:'',model:'',keyUrl:'',docUrl:''}
];
const byId=id=>PROVIDERS.find(p=>p.id===id);
export function checkBaseUrl(value){
  let u;try{u=new URL(String(value||'').trim());}catch{throw Error('接口地址格式不对，应以 https:// 开头。');}
  const local=/^(127\.0\.0\.1|localhost|\[::1\])$/.test(u.hostname);
  if(u.protocol!=='https:'&&!(local&&u.protocol==='http:'))throw Error('为保护密钥，接口地址必须是 https://（本机 Ollama 除外）。');
  if(u.username||u.password||u.search||u.hash)throw Error('接口地址里不要带账号、参数或 # 号。');
  return u.origin+u.pathname.replace(/\/+$/,'');
}
// 设置文件 ai.json：apiKey 在 Windows 上是 DPAPI 加密后的内容（keyProtection='dpapi'），keyHint 只存打码提示。
function normalizeConfig(input,old={}){
  const p=byId(String(input.provider||''));if(!p)throw Error('请选择 AI 服务商。');
  const baseUrl=checkBaseUrl(input.baseUrl||p.baseUrl),model=String(input.model||'').trim();
  if(model.length>120||/[\s\u0000-\u001f]/.test(model))throw Error('模型名不能含空格，最多 120 个字符。');
  const mode=input.privacy===undefined?(old.privacy&&old.privacy.mode)||'mask':String(input.privacy&&input.privacy.mode||input.privacy);
  if(!PRIVACY_MODES.includes(mode))throw Error('隐私保护方式不对。');
  let key=null;// null = 不改动已保存的 Key
  if(input.clearKey)key='';else if(input.apiKey!==undefined&&String(input.apiKey).trim())key=String(input.apiKey).trim();
  if(key&&(key.length>400||/[\s\u0000-\u001f]/.test(key)))throw Error('API Key 格式不对：不能有空格或换行，请重新复制。');
  return {provider:p.id,baseUrl,model,key,privacy:{mode}};
}
export function maskKey(k){k=String(k||'');return k?(k.length<=8?'****':k.slice(0,3)+'…'+k.slice(-4)):'';}
export function createAiStore(file,{secrets=createSecretBox()}={}){
  async function read(){try{const d=JSON.parse(await fs.readFile(file,'utf8'));return d&&typeof d==='object'?d:{};}catch{return {};}}
  async function write(d){await fs.mkdir(path.dirname(file),{recursive:true});const tmp=file+'.'+process.pid+'.'+Date.now()+'.tmp';await fs.writeFile(tmp,JSON.stringify(d,null,2),'utf8');await fs.rename(tmp,file);}
  async function seal(plain){try{const r=await secrets.protect(plain);return {apiKey:r.value,keyProtection:r.protection,keyWarning:secrets.available&&r.protection!=='dpapi'&&r.protection!=='keychain'?'加密失败，Key 以明文保存':''};}
    catch{return {apiKey:plain,keyProtection:'none',keyWarning:'系统加密不可用，Key 以明文保存在本机'};}}
  // 旧版本（0.10.x）明文保存的 Key：在 Windows 上读到时顺便加密。
  async function migrate(d){if(d.apiKey&&d.keyProtection!=='dpapi'&&d.keyProtection!=='keychain'&&secrets.available){const sealed=await seal(d.apiKey);if(sealed.keyProtection==='dpapi'||sealed.keyProtection==='keychain'){const next={...d,...sealed,keyHint:d.keyHint||maskKey(d.apiKey)};await write(next);return next;}}return d;}
  return {
    async raw(){return read();},
    async mode(){const d=await read();return (d.privacy&&d.privacy.mode)||'mask';},
    async get(){let d=await migrate(await read());const apiKey=d.apiKey?await secrets.unprotect(d.apiKey,d.keyProtection):'';return {...d,apiKey};},
    async publicView(){const c=await read(),p=byId(c.provider);const hint=c.keyHint||(c.keyProtection!=='dpapi'&&c.keyProtection!=='keychain'?maskKey(c.apiKey):'');
      return {providers:PROVIDERS,config:{provider:c.provider||'',baseUrl:c.baseUrl||'',model:c.model||'',hasKey:!!c.apiKey,keyHint:hint,keyProtection:c.apiKey?(c.keyProtection==='dpapi'||c.keyProtection==='keychain'?c.keyProtection:'none'):'',keyWarning:c.keyWarning||'',privacy:{mode:(c.privacy&&c.privacy.mode)||'mask'},ready:!!(p&&c.baseUrl&&c.model&&(c.apiKey||p.noKey))}};},
    async save(input){const old=await read();
      // “不使用 AI”：清空服务商和模型（已保存的 Key 留着，以后重新选回来不用再填）。
      // 选“不使用 AI”时也可以顺便清除 Key、改个人信息保护方式。
      if(input&&input.provider==='none'){const mode=PRIVACY_MODES.includes(input.privacy&&input.privacy.mode)?input.privacy.mode:null;
        await write({...old,provider:'',baseUrl:'',model:'',...(input.clearKey?{apiKey:'',keyProtection:'none',keyHint:'',keyWarning:''}:{}),...(mode?{privacy:{...(old.privacy||{}),mode}}:{}),updated:new Date().toISOString()});return this.publicView();}
      const n=normalizeConfig(input||{},old);let keyPart;
      if(n.key===null)keyPart={apiKey:old.apiKey||'',keyProtection:old.keyProtection||'none',keyHint:old.keyHint||'',keyWarning:old.keyWarning||''};
      else if(n.key==='')keyPart={apiKey:'',keyProtection:'none',keyHint:'',keyWarning:''};
      else keyPart={...await seal(n.key),keyHint:maskKey(n.key)};
      await write({provider:n.provider,baseUrl:n.baseUrl,model:n.model,...keyPart,privacy:n.privacy,updated:new Date().toISOString()});return this.publicView();}
  };
}
function friendly(status,body){
  const raw=String(body&&(body.error?.message||body.error||body.message)||'').slice(0,200);
  if(status===401||status===403)return 'API Key 不正确、已失效或没有权限，请到官网重新创建后粘贴。';
  if(status===402)return '账户余额不足，请到服务商官网充值或领取免费额度。';
  if(status===429)return '请求太频繁或额度用完了，请稍后再试，或到官网查看额度。';
  if(status===404)return '找不到这个模型或接口地址，请点“获取模型列表”重新选择模型。'+(raw?'（'+raw+'）':'');
  if(status===400)return '服务商拒绝了请求：'+(raw||'请检查模型名。');
  return 'AI 服务返回错误（'+status+'）'+(raw?'：'+raw:'');
}
// 只调用用户保存的地址；不带 temperature（部分新模型只接受默认值）。
export function createAiClient({store,fetchImpl=fetch,timeout=120000,guard=null}){
  async function config(){const c=await store.get(),p=byId(c.provider);if(!p||!c.baseUrl)throw Object.assign(Error('还没有设置 AI。请点上方“设置”，在“AI 服务”里接入自己的 AI（设置页里有一步步获取 API Key 的说明）。'),{status:409});
    if(!c.apiKey&&!p.noKey)throw Object.assign(Error('还没有填写 API Key。不会申请的话，点设置页里的“申请教程”。'),{status:409});return {...c,p};}
  function headers(c){const h={'Content-Type':'application/json',Accept:'application/json'};if(c.apiKey)h.Authorization='Bearer '+c.apiKey;return h;}
  // 本机 AI（Ollama 等）：正式请求前用 3 秒确认它在运行，没开时立刻说明，不让用户干等。
  async function localReady(c){try{await fetchImpl(c.baseUrl+'/models',{method:'GET',headers:headers(c),signal:AbortSignal.timeout(3000)});}catch(e){throw Object.assign(Error('本机 AI（'+(c.p.name||'Ollama').replace(/（.*$/,'')+'）没有运行：请先打开它，或到“设置 → AI 服务”换成别的服务商、或选“不使用 AI”。'),{status:503});}}
  async function call(url,init,c){if(c.p&&c.p.region==='local'&&init.method==='POST')await localReady(c);let r;try{r=await fetchImpl(url,{...init,headers:headers(c),signal:AbortSignal.timeout(c.p&&c.p.region==='local'?Math.max(timeout,300000):timeout)});}
    catch(e){throw Object.assign(Error(e.name==='TimeoutError'?'AI 服务响应超时，请稍后重试或换一个更快的模型。':'连不上 AI 服务：请检查网络、公司代理或接口地址。'),{status:502});}
    const text=await r.text();let body=null;try{body=JSON.parse(text);}catch{}
    if(!r.ok)throw Object.assign(Error(friendly(r.status,body)),{status:502,upstream:r.status});
    if(!body)throw Object.assign(Error('AI 服务返回的不是有效数据，请检查接口地址。'),{status:502});return body;}
  async function models(){const c=await config(),b=await call(c.baseUrl+'/models',{method:'GET'},c);const list=(b.data||b.models||[]).map(m=>String(m.id||m.name||'').replace(/^models\//,'')).filter(Boolean);return [...new Set(list)].sort().slice(0,300);}
  async function chat(messages,{maxTokens=4000,feature='',confirmed=false}={}){const c=await config();if(!c.model)throw Object.assign(Error('还没有选择模型，请点“获取模型列表”后选一个。'),{status:409});
    const g=guard?await guard.beforeSend({messages,feature,confirmed,provider:c.provider,region:c.p.region,model:c.model}):{messages,session:null,report:null};
    const b=await call(c.baseUrl+'/chat/completions',{method:'POST',body:JSON.stringify({model:c.model,messages:g.messages,max_tokens:maxTokens,stream:false})},c);
    const msg=b.choices&&b.choices[0]&&b.choices[0].message,content=typeof msg?.content==='string'?msg.content:Array.isArray(msg?.content)?msg.content.map(x=>x.text||'').join(''):'';
    const answer=content.replace(/<think>[\s\S]*?<\/think>/g,'').replace(/^[\s\S]*?<\/think>/,'').trim();
    // 思考型模型（如 deepseek-r1）可能把额度都用在思考上，正文为空
    if(!answer)throw Object.assign(Error(msg&&(msg.reasoning||msg.reasoning_content)||b.choices?.[0]?.finish_reason==='length'?'模型把字数额度都用在“思考”上了，没有给出正文。请换一个非思考模型，或稍后重试。':'模型没有返回内容，请换一个模型重试。'),{status:502,empty:true});
    const restored=g.session?await guard.afterReceive(g.session,answer):answer;return {answer:restored,model:c.model,provider:c.provider,privacy:g.report};}
  // 测试连接：服务返回了有效答复就算连通；思考型模型可能只给出思考内容，也算连通。
  async function test(){const t=Date.now();try{const r=await chat([{role:'user',content:'请只回复：OK'}],{maxTokens:512,feature:'test'});return {ok:true,model:r.model,ms:Date.now()-t,reply:r.answer.slice(0,40)};}
    catch(e){if(e.privacy)return {ok:false,error:e.message,models:[]};if(e.empty){const c=await store.get();return {ok:true,model:c.model,ms:Date.now()-t,reply:'',thinking:true};}let list=[];if(e.upstream===404||e.upstream===400){try{list=await models();}catch{}}return {ok:false,error:e.message,models:list};}}
  return {chat,models,test,config};
}
// ---------- AI 生成 PPT 大纲 ----------
const LAYOUT_TAG={zh:{bullets:'要点',grid:'卡片',stat:'数据',timeline:'时间线',quote:'引言',compare:'对比'},ja:{bullets:'箇条書き',grid:'カード',stat:'数値',timeline:'タイムライン',quote:'引用',compare:'比較'}};
export function outlinePrompt({topic,audience,pages,lang,notes}){
  const ja=lang==='ja',n=Math.max(3,Math.min(12,Number(pages)||6));
  const sys=ja?'あなたはプレゼン資料の構成を考えるプロの編集者です。事実を誇張せず、具体的で短い表現にしてください。出力は JSON のみ。'
    :'你是专业的演示文稿策划编辑。内容要具体、简短、不夸大事实，不编造具体数据（没有把握的数字用“待补充”）。只输出 JSON。';
  const user=(ja?`テーマ：${topic}\n対象：${audience||'一般'}\n章の数：約 ${n}\n補足：${notes||'なし'}\n`:`主题：${topic}\n听众：${audience||'一般听众'}\n章节数：约 ${n} 个\n补充要求：${notes||'无'}\n`)+
    `JSON 格式：{"title":"...","subtitle":"...","sections":[{"heading":"...","layout":"bullets|grid|stat|timeline|quote|compare","points":["..."],"notes":"..."}]}\n`+
    (ja?'ルール：各章 2〜5 項目、1 項目 40 文字以内。grid は「見出し：説明」形式で 3〜6 項目。stat は数字で始まる 2〜4 項目。timeline は日付や段階で始まる 3〜6 項目。quote は 1 項目。compare は前半が比較元・後半が比較先の偶数項目で、最初の行を「A｜B」形式の見出しにする。レイアウトを偏らせない。notes は発表者が話す内容を 2〜3 文で。日本語で書く。'
      :'规则：每章 2～5 条，每条不超过 40 字。grid 用“小标题：说明”格式，3～6 条；stat 每条以数字开头，2～4 条；timeline 每条以时间或阶段开头，3～6 条；quote 只有 1 条；compare 是左右对比，第一条写“左边标题｜右边标题”，其余前一半是左边、后一半是右边。各章版式要有变化。notes 写演讲者要说的话，2～3 句。用简体中文。');
  return [{role:'system',content:sys},{role:'user',content:user}];
}
export function parseOutline(answer,lang){
  const s=String(answer||'');const a=s.indexOf('{'),b=s.lastIndexOf('}');if(a<0||b<a)throw Error('AI 没有按格式返回大纲，请再试一次。');
  let d;try{d=JSON.parse(s.slice(a,b+1));}catch{throw Error('AI 返回的大纲格式有误，请再试一次。');}
  const tags=LAYOUT_TAG[lang==='ja'?'ja':'zh'],one=v=>String(v??'').replace(/[\r\n]+/g,' ').trim();
  const sections=(Array.isArray(d.sections)?d.sections:[]).slice(0,20).map(x=>({heading:one(x.heading).slice(0,90),layout:tags[x.layout]?x.layout:'',points:(Array.isArray(x.points)?x.points:[]).map(one).filter(Boolean).slice(0,12).map(p=>p.slice(0,300).replace(/^>\s*/,'')),notes:one(x.notes).slice(0,600)})).filter(x=>x.heading&&x.points.length);
  if(!sections.length)throw Error('AI 返回的大纲是空的，请换个说法再试。');
  return {title:one(d.title).slice(0,80)||sections[0].heading.slice(0,80),subtitle:one(d.subtitle).slice(0,150),outline:sections.map(x=>x.heading+(x.layout?` [${tags[x.layout]}]`:'')+'\n'+x.points.join('\n')+(x.notes?'\n> '+x.notes:'')).join('\n\n')};
}
