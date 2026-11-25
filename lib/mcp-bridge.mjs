import {randomUUID} from 'node:crypto';
const fault=(message,status=422)=>Object.assign(Error(message),{status});
export function createMcpBridge({timeout=30000,now=Date.now}={}){
 const clients=new Map(),logs=[];let active=null;
 const log=(action,status)=>{logs.push({at:new Date(now()).toISOString(),action,status});if(logs.length>60)logs.shift();};
 function live(){return [...clients.values()].filter(c=>now()-c.at<6000);}
 function register(input){if(!/^[a-f0-9]{16}$/.test(input.clientId)||!['wps','et','wpp'].includes(input.host)||!['wps','office'].includes(input.suite))throw fault('连接信息无效。');const tools=[].concat(input.tools||[]).slice(0,500).filter(x=>/^[a-zA-Z][a-zA-Z0-9]{0,79}$/.test(x.id)).map(x=>({id:x.id,label:String(x.label||x.id).slice(0,100),desc:String(x.desc||'').slice(0,300),fields:[].concat(x.fields||[]).slice(0,30).filter(f=>/^[a-zA-Z][a-zA-Z0-9]{0,49}$/.test(f.key)).map(f=>({key:f.key,label:String(f.label||f.key),required:f.required!==false}))}));clients.set(input.clientId,{clientId:input.clientId,host:input.host,suite:input.suite,tools,at:now()});return {ok:true};}
 function poll(input){const c=clients.get(input.clientId);if(!c)throw fault('请重新连接工具箱。',409);c.at=now();if(!active||active.clientId!==c.clientId||active.delivered)return {task:null};active.delivered=true;return {task:{id:active.id,action:active.action,args:active.args,deadline:active.deadline}};}
 function finish(input){if(!active||input.id!==active.id||input.clientId!==active.clientId)throw fault('任务已结束。',409);const task=active;const c=clients.get(input.clientId);if(c)c.at=now();active=null;clearTimeout(task.timer);if(!task.expired){if(input.error){task.reject(fault(String(input.error).slice(0,1000)));log(task.action,'failed');}else{task.resolve(input.result);log(task.action,'complete');}}return {ok:true};}
 function execute(input){if(active)throw fault('上一个任务尚未结束，请在文档中处理确认提示。',409);const c=live().find(c=>c.clientId===input.clientId);if(!c)throw fault('目标窗口未连接，请打开工具箱侧栏。',409);if(!['readSelection','readDocument','insertText','runTool'].includes(input.action))throw fault('不支持这个操作。');const args=input.args||{};
  if(input.action==='insertText'&&(typeof args.text!=='string'||!args.text.trim()||args.text.length>50000))throw fault('插入内容应为 1～50000 个字符。');
  if(input.action==='runTool'){const tool=c.tools.find(t=>t.id===args.id);if(!tool)throw fault('工具未登记。');const options=args.options||{};if(typeof options!=='object'||Array.isArray(options)||Object.keys(options).some(k=>!tool.fields.some(f=>f.key===k)||['__proto__','constructor','prototype'].includes(k))||Object.values(options).some(v=>typeof v!=='string'||v.length>10000))throw fault('工具参数无效。');for(const f of tool.fields)if(f.required&&!String(options[f.key]||'').trim())throw fault('请填写：'+f.label);}
  return new Promise((resolve,reject)=>{const task={id:randomUUID(),clientId:c.clientId,action:input.action,args,deadline:now()+timeout,resolve,reject,delivered:false,expired:false};task.timer=setTimeout(()=>{task.expired=true;task.reject(fault('任务超过 30 秒。请在文档中关闭确认提示后重试。',408));log(task.action,'timeout');if(!task.delivered)active=null;},timeout);active=task;log(task.action,'started');});
 }
 function close(){if(active){clearTimeout(active.timer);if(!active.expired)active.reject(fault('服务已停止。',503));active=null;}}
 return {register,poll,finish,execute,close,status:()=>({clients:live(),busy:!!active,logs:[...logs]})};
}
