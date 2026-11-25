// 前端回归测试用的最小替身：假时钟、假 DOM、假 WPS Application、假 fetch。只实现被测脚本实际用到的部分。
import vm from 'node:vm';
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const PUBLIC=path.join(path.dirname(fileURLToPath(import.meta.url)),'..','..','public');
export const flush=async()=>{for(let i=0;i<8;i++)await new Promise(resolve=>setImmediate(resolve));};

export function createClock(){
  let now=0,seq=0;const timers=new Map();
  const add=(fn,ms,every)=>{const id=++seq;timers.set(id,{fn,at:now+(ms||0),every});return id;};
  return {
    setTimeout:(fn,ms)=>add(fn,ms,0),setInterval:(fn,ms)=>add(fn,ms,ms),
    clearTimeout:id=>{timers.delete(id);},clearInterval:id=>{timers.delete(id);},
    async advance(ms){
      const end=now+ms;
      for(;;){
        let next=null,nextId=0;for(const [id,t] of timers)if(t.at<=end&&(!next||t.at<next.at)){next=t;nextId=id;}
        if(!next)break;now=next.at;if(next.every)next.at+=next.every;else timers.delete(nextId);next.fn();await flush();
      }
      now=end;await flush();
    }
  };
}

class FakeElement {
  constructor(doc,id,tag='div'){this.ownerDocument=doc;this.id=id;this.tagName=tag;this.children=[];this.attrs={};this.listeners={};this.dataset={};this.value='';this.checked=false;this.disabled=false;this.hidden=false;this.textContent='';this.removed=false;this.style={};this._classes=new Set();this.found={};
    this.classList={toggle:(name,on)=>{(on===undefined?!this._classes.has(name):on)?this._classes.add(name):this._classes.delete(name);},add:name=>this._classes.add(name),remove:name=>this._classes.delete(name),contains:name=>this._classes.has(name)};}
  get className(){return [...this._classes].join(' ');}
  set className(value){this._classes=new Set(String(value).split(/\s+/).filter(Boolean));}
  append(...nodes){this.children.push(...nodes);}
  replaceChildren(...nodes){this.children=[...nodes];}
  setAttribute(name,value){this.attrs[name]=String(value);}
  getAttribute(name){return this.attrs[name];}
  addEventListener(type,fn){(this.listeners[type]??=[]).push(fn);}
  removeEventListener(){}
  dispatch(type,event={}){(this.listeners[type]||[]).forEach(fn=>fn(event));if(typeof this['on'+type]==='function')this['on'+type](event);}
  click(){this.dispatch('click');}
  focus(){}
  remove(){this.removed=true;}
  querySelector(selector){return this.found[selector]||null;}
}
export function createDocument(){
  const byId=new Map(),groups={},doc={created:[],body:null};
  doc.getElementById=id=>{if(!byId.has(id))byId.set(id,new FakeElement(doc,id));return byId.get(id);};
  doc.createElement=tag=>{const el=new FakeElement(doc,'',tag);doc.created.push(el);return el;};
  doc.createElementNS=(ns,tag)=>new FakeElement(doc,'',tag);
  doc.querySelectorAll=selector=>groups[selector]||[];
  doc.querySelector=selector=>(groups[selector]||[])[0]||null;
  doc.group=(selector,elements)=>{groups[selector]=elements;return elements;};
  doc.addEventListener=()=>{};
  doc.body=new FakeElement(doc,'body','body');
  return doc;
}

export function createApp(){
  const panes=[],store=new Map();
  return {panes,dialogs:[],
    PluginStorage:{getItem:key=>store.has(key)?store.get(key):null,setItem:(key,value)=>{store.set(key,String(value));}},
    CreateTaskPane(url,title){const pane={ID:panes.length+1,url,title,Visible:false,Width:0,deleted:false,Delete(){this.deleted=true;}};panes.push(pane);return pane;},
    GetTaskPane(id){return panes.find(pane=>pane.ID===id&&!pane.deleted)||null;},
    ShowDialog(...args){this.dialogs.push(args);}};
}

// 同源的两个脚本环境（功能区、侧栏）共享的“浏览器存储与广播”。
export function createSharedBrowser({broadcast=true,localStorage=true}={}){
  const storage=new Map(),channels=new Map();
  return {
    localStorage:localStorage?{getItem:key=>storage.has(key)?storage.get(key):null,setItem:(key,value)=>{storage.set(key,String(value));}}:undefined,
    BroadcastChannel:broadcast?class{
      constructor(name){this.name=name;this.onmessage=null;if(!channels.has(name))channels.set(name,new Set());channels.get(name).add(this);}
      postMessage(data){for(const peer of channels.get(this.name))if(peer!==this)queueMicrotask(()=>peer.onmessage&&peer.onmessage({data}));}
      close(){channels.get(this.name).delete(this);}
    }:undefined
  };
}

export function loadScripts(sandbox,files){
  sandbox.window=sandbox;vm.createContext(sandbox);
  for(const file of files)new vm.Script(fs.readFileSync(path.join(PUBLIC,file),'utf8'),{filename:file}).runInContext(sandbox);
  return sandbox;
}

export function createFetch(){
  const calls=[];let token=0;
  const reply=(status,body)=>({ok:status>=200&&status<300,status,json:async()=>{if(body===undefined)throw SyntaxError('not json');return body;},blob:async()=>body});
  const fetch=(url,options={})=>new Promise((resolve,reject)=>{
    const call={url,options,body:options.body?JSON.parse(options.body):null,answered:false,respond:(status,body)=>{call.answered=true;resolve(reply(status,body));}};
    const abort=()=>reject(Object.assign(Error('aborted'),{name:'AbortError'}));
    if(options.signal){if(options.signal.aborted)return abort();options.signal.addEventListener('abort',abort);}
    calls.push(call);
    if(url==='/api/session')call.respond(200,{token:'token-'+(++token)});
  });
  return {fetch,calls,studio:action=>calls.filter(call=>call.url==='/api/studio/'+action)};
}
