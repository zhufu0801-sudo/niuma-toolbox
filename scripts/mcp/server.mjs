// MCP stdio：stdout 仅输出一行一个 JSON-RPC 消息。兼容 2024-11-05 / 2025-03-26 / 2025-06-18 / 2025-11-25。
import {createInterface} from 'node:readline';
import {randomUUID} from 'node:crypto';
const sessionId=randomUUID();
const maskNote='默认将个人信息换成代号，写回时请保留代号，工具箱会还原；未知代号原样保留。個人情報は既定で記号化されます。書き戻すときは記号を残してください。対応表にある記号を元に戻します。';
const base='http://127.0.0.1:39127';let token='',initialized=false;const versions=['2024-11-05','2025-03-26','2025-06-18','2025-11-25'];
async function api(action,input={}){if(!token){const r=await fetch(base+'/api/session',{signal:AbortSignal.timeout(5000)});if(!r.ok)throw Error('本机服务未连接。');token=(await r.json()).token;}const r=await fetch(base+'/api/mcp/'+action,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify({...input,sessionId}),signal:AbortSignal.timeout(35000)});const d=await r.json();if(!r.ok){if(r.status===403)token='';throw Error(d.error||'操作失败。');}return d;}
const target={clientId:{type:'string',description:'从 list_windows 获取的窗口 ID'}};
const schema=(properties,required)=>({type:'object',properties,required,additionalProperties:false});
async function tools(){const state=await api('status');return [
 {name:'list_windows',description:'列出已连接的 WPS / Office 窗口和可用工具。Office 侧栏须保持打开。',inputSchema:schema({},[])},
 {name:'read_selection',description:'读取选中内容（最多 20000 字符）。'+maskNote,inputSchema:schema(target,['clientId'])},
 {name:'read_document',description:'读取当前文档文本（最多 20000 字符）。'+maskNote,inputSchema:schema(target,['clientId'])},
 {name:'insert_text',description:'经用户在文档中确认后插入文字。'+maskNote,inputSchema:schema({...target,text:{type:'string',maxLength:50000}},['clientId','text'])},
 ...state.clients.flatMap(c=>c.tools.map(t=>({name:'run_'+c.clientId+'_'+t.id,description:c.suite+' '+c.host+'：'+t.label+'；'+t.desc+'。'+maskNote,inputSchema:schema(Object.fromEntries(t.fields.map(f=>[f.key,{type:'string',description:f.label}])),t.fields.filter(f=>f.required).map(f=>f.key))})))];}
async function call(name,args){if(name==='list_windows')return api('status');if(name==='read_selection'||name==='read_document'||name==='insert_text')return api('execute',{clientId:args.clientId,action:{read_selection:'readSelection',read_document:'readDocument',insert_text:'insertText'}[name],args:name==='insert_text'?{text:args.text}:{}});const m=/^run_([a-f0-9]{16})_([A-Za-z][A-Za-z0-9]{0,79})$/.exec(name);if(!m)throw Error('工具未登记。');return api('execute',{clientId:m[1],action:'runTool',args:{id:m[2],options:args}});}
const send=x=>process.stdout.write(JSON.stringify(x)+'\n');
async function message(line){let q;try{q=JSON.parse(line);}catch{return send({jsonrpc:'2.0',id:null,error:{code:-32700,message:'Parse error'}});}if(!q||q.jsonrpc!=='2.0'||typeof q.method!=='string')return send({jsonrpc:'2.0',id:q?.id??null,error:{code:-32600,message:'Invalid request'}});if(q.id===undefined){if(q.method==='notifications/initialized')initialized=true;return;}
 try{let result;if(q.method==='initialize'){initialized=true;result={protocolVersion:versions.includes(q.params?.protocolVersion)?q.params.protocolVersion:'2025-11-25',capabilities:{tools:{}},serverInfo:{name:'office-toolbox',version:'0.16.0'}};}
 else if(q.method==='ping')result={};else if(!initialized)throw Object.assign(Error('Initialize first'),{code:-32002});else if(q.method==='tools/list')result={tools:await tools()};else if(q.method==='tools/call'){try{const value=await call(q.params?.name,q.params?.arguments||{});result={content:[{type:'text',text:JSON.stringify(value)}]};}catch(e){result={isError:true,content:[{type:'text',text:e.message}]};}}else throw Object.assign(Error('Method not found'),{code:-32601});send({jsonrpc:'2.0',id:q.id,result});}catch(e){send({jsonrpc:'2.0',id:q.id,error:{code:e.code||-32603,message:e.message}});}}
let busy=false;const input=createInterface({input:process.stdin,crlfDelay:Infinity});input.on('line',line=>{if(line.length>300000)return send({jsonrpc:'2.0',id:null,error:{code:-32600,message:'Message too large'}});message(line).catch(e=>process.stderr.write(e.message+'\n'));});
