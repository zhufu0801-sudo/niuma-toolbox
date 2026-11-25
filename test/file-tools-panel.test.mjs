import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const src=fs.readFileSync(new URL('../public/file-tools-panel.js',import.meta.url),'utf8');
function load({app,office,reply}={}){
  const stub=()=>({value:'',textContent:'',hidden:false,disabled:false,className:'',dataset:{},children:[],append(...x){this.children.push(...x);},addEventListener(){},click(){this.clicked=true;},focus(){},select(){},scrollIntoView(){}});
  const els={},buttons=[];const section={className:'',dataset:{},set innerHTML(h){for(const m of h.matchAll(/<(\w+)[^>]*id="ft([A-Za-z]+)"/g)){const e=stub();els[m[2]]=e;if(m[1]==='button')buttons.push(e);}},querySelector:s=>s==='details'?{open:false}:els[s.slice(3)],querySelectorAll:()=>buttons,addEventListener(){}};
  const calls=[];const w={OfficeTools:{lib:{ask:async()=>true},app:()=>app,host:()=>'wps'},AiTextPanel:{host:()=>({insert:async t=>'已插入：'+t})},
    fetch:async(url,o)=>{if(url==='/api/session')return {json:async()=>({token:'t'})};const body=JSON.parse(o.body);calls.push({url,body});const d=reply?reply(url,body):{};return {ok:!d.error,json:async()=>d};},
    document:{createElement:t=>t==='section'?section:stub(),getElementById:id=>id==='aiChatPanel'?{parentNode:{insertBefore(){}},nextSibling:null}:null,addEventListener(){},execCommand:()=>true},navigator:{}};
  if(office){w.AiTextHost={};w.Office={context:{document:{url:office}}};}
  w.window=w;vm.runInNewContext(src,w);return {w,els,calls};
}
const flush=()=>new Promise(r=>setTimeout(r,0));
test('Office 图片识别与历史记录都使用刚保存后的文件路径',async()=>{
 const {w,els,calls}=load({office:'C:/old.docx',reply:()=>({text:'测试',count:1,engine:'Windows',items:[]})});
 let n=0;w.OfficeAddinTools={documentUrl:async()=>{n++;return 'file:///C:/new%20file.docx';},localPath:()=> 'C:/new file.docx'};
 els.DocOcr.onclick();await flush();await flush();els.History.onclick();await flush();await flush();
 assert.equal(n,2);assert.equal(calls.length,2);assert.ok(calls.every(c=>c.body.path==='C:/new file.docx'));
});
test('文件与图片：提取 / 压缩用当前文档的路径；未保存的文档先提示',async()=>{
  const doc={FullName:'C:\\资料\\报告.docx',Saved:true},app={ActiveDocument:doc,Documents:{Open(p){app.opened=p;}}};
  const {els,calls}=load({app,reply:url=>url.endsWith('extract-images')?{dir:'C:\\资料\\报告-图片',count:3}:url.endsWith('compress-images')?{changed:2,count:3,before:3e6,after:1e6,fileBefore:4e6,fileAfter:2e6,path:'C:\\资料\\报告-压缩.docx'}:{}});
  els.Extract.onclick();await flush();await flush();assert.deepEqual(calls[0],{url:'/api/tools/extract-images',body:{path:'C:\\资料\\报告.docx'}});assert.match(els.Status.textContent,/已把 3 张图片存到/);assert.equal(els.Result.hidden,false);
  els.Level.value='60,1280';els.Compress.onclick();await flush();await flush();assert.deepEqual(calls[1].body,{path:'C:\\资料\\报告.docx',quality:60,maxSide:1280});assert.match(els.Status.textContent,/2.9 MB → 977 KB，文件 3.8 MB → 1.9 MB/);
  els.Open.onclick();assert.equal(app.opened,'C:\\资料\\报告-压缩.docx');
  const u=load({app:{ActiveDocument:{FullName:'文档1',Saved:false}}});u.els.Extract.onclick();await flush();await flush();assert.equal(u.calls.length,0);assert.match(u.els.Status.textContent,/还没有保存到电脑上/);
});
test('文件与图片：识别结果放进文本框；Office 里用文档地址、网上的文档提示先另存',async()=>{
  const {w,els,calls}=load({office:'C:\\a\\b.pptx',reply:()=>({text:'会议室',count:2,engine:'Windows',note:''})});
  els.DocOcr.onclick();await flush();await flush();await flush();assert.deepEqual(calls[0].body,{path:'C:\\a\\b.pptx',lang:'zh'});assert.equal(els.Text.value,'会议室');assert.equal(els.Text.hidden,false);
  const n=load({office:'https://x.sharepoint.com/b.docx'});n.els.History.onclick();await flush();await flush();assert.match(n.els.Status.textContent,/OneDrive/);
  assert.equal(w.FileToolsPanel.select('pdf'),true);assert.equal(w.FileToolsPanel.select('nope'),false);
});
test('服务重启后旧口令失效：自动重新取口令再试一次，不用重开 WPS',async()=>{
  const doc={FullName:'C:\\资料\\报告.docx',Saved:true},app={ActiveDocument:doc,Documents:{Open(){}}};
  const {w,els}=load({app});let n=0;const sent=[];
  w.fetch=async(url,o)=>{if(url==='/api/session'){n++;return {json:async()=>({token:'t'+n})};}const tk=o.headers['X-Coach-Token'];sent.push(tk);
    if(tk!=='t2')return {ok:false,status:403,json:async()=>({error:'本机连接已过期，请重新打开工具箱。'})};return {ok:true,status:200,json:async()=>({dir:'C:\\资料\\报告-图片',count:1})};};
  els.Extract.onclick();await flush();await flush();await flush();await flush();
  assert.deepEqual(sent,['t1','t2']);assert.match(els.Status.textContent,/1 张图片/);
  // 新口令也被拒绝时只重试一次，然后把服务的提示告诉用户
  w.fetch=async(url)=>url==='/api/session'?{json:async()=>({token:'bad'})}:{ok:false,status:403,json:async()=>({error:'本机连接已过期，请重新打开工具箱。'})};
  els.Extract.onclick();await flush();await flush();await flush();await flush();assert.match(els.Status.textContent,/过期/);
});
