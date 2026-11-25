// 粘贴栏（2026-11-21）：存储、敏感内容过滤、服务接口、剪贴板监听、插入工具。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import os from 'node:os';import path from 'node:path';import vm from 'node:vm';import {EventEmitter} from 'node:events';
import {createPasteBar,sensitiveReason,blockFromText} from '../lib/pastebar.mjs';
import {watchClipboard} from '../lib/clipboard-watch.mjs';
import {validateSettings,defaultSettings} from '../lib/settings.mjs';
import {createCoachServer} from '../server.mjs';
import {toolsCode} from './helpers/tools-code.mjs';

test('粘贴栏：像密码、身份证号、银行卡号、マイナンバー、API Key 的内容不自动记录；普通文字、电话、网址照常',()=>{
  for(const s of ['Abcd1234!','P@ssw0rd2026','110101199003071234','4111 1111 1111 1111','1234 5678 9012','sk-abcdefghijklmnop1234','ghp_abcdefghijklmnop1234567890','AKIAABCDEFGHIJKLMNOP','dGhpc2lzYXZlcnlsb25nYmFzZTY0c3RyaW5nMTIzNA=='])assert.ok(sensitiveReason(s),s);
  for(const s of ['本合同总金额为人民币 128560.50 元','13800000001','03-1234-5678','https://example.com/a?b=1','pat@example.com','会議は 11 月 20 日です','2026年11月20日','A4 复印纸\t10 箱','report'])assert.equal(sensitiveReason(s),null,s);
  assert.equal(sensitiveReason('Report2026'),'password','大小写字母 + 数字的单个词宁可当密码跳过');
});
test('粘贴栏：从表格复制的“制表符 + 换行”存成表格，其余存成文字',()=>{
  assert.deepEqual(blockFromText('姓名\t电话\r\n张三\t138\r\n'),{type:'table',rows:[['姓名','电话'],['张三','138']]});
  assert.deepEqual(blockFromText('第一行\n第二行'),{type:'text',text:'第一行\n第二行'});
  assert.equal(blockFromText('a\tb\nc').type,'text','列数不齐按文字');
});
test('粘贴栏：新的在前；重复内容移到最前不重复存；普通最多 30 条，钉住的另算且排最前；清空保留钉住的',async()=>{
  let now=1000;const bar=createPasteBar(null,{limit:3,now:()=>now++});
  const a=await bar.add('甲','clipboard','clipboard');await bar.add('乙','wps');await bar.add({type:'text',text:'丙'},'wps');
  assert.deepEqual((await bar.list()).map(x=>x.block.text),['丙','乙','甲']);
  await bar.add('甲','clipboard','clipboard');assert.deepEqual((await bar.list()).map(x=>x.block.text),['甲','丙','乙'],'重复内容移到最前');
  await bar.pin(a.id,true);await bar.add('丁','wps');await bar.add('戊','wps');await bar.add('己','wps');
  const list=await bar.list();assert.equal(list[0].block.text,'甲');assert.ok(list[0].pinned);assert.deepEqual(list.slice(1).map(x=>x.block.text),['己','戊','丁'],'普通的只留 3 条，钉住的不算');
  assert.equal(await bar.clear(),1);assert.deepEqual((await bar.list()).map(x=>x.block.text),['甲']);
  await assert.rejects(bar.add('x'.repeat(10001),'wps'),/最多 1 万字/);await assert.rejects(bar.add('a','drop'),/来源无效/);
});
test('粘贴栏：保存在文件里，重启服务后还在',async()=>{
  const dir=fs.mkdtempSync(path.join(os.tmpdir(),'pb-')),file=path.join(dir,'pastebar.json');
  await createPasteBar(file).add('留着','wps');assert.equal((await createPasteBar(file).list())[0].block.text,'留着');
});
test('设置：粘贴栏自动记录默认关闭，只接受开 / 关',()=>{
  assert.equal(defaultSettings('/b').pasteBar.watch,false);
  assert.equal(validateSettings({pasteBar:{watch:true}},defaultSettings('/b')).pasteBar.watch,true);
  assert.throws(()=>validateSettings({pasteBar:{watch:'yes'}},defaultSettings('/b')),/粘贴栏选项/);
});

async function serve(t,options={}){
  const server=createCoachServer(options);await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await (await fetch(base+'/api/session')).json()).token;
  const call=(p,method='GET',body)=>fetch(base+p,{method,headers:{'Content-Type':'application/json','X-Coach-Token':token},body:body?JSON.stringify(body):undefined}).then(async r=>({status:r.status,data:await r.json()}));
  return {base,call};
}
test('服务：打开“自动记录”才开始监听剪贴板，关掉就停；敏感内容跳过并计数；需要本机口令',async t=>{
  let started=0,stopped=0,emit=null;
  const watcher=on=>{started++;emit=on;return {stop(){stopped++;}};};
  const {base,call}=await serve(t,{clipboardWatcher:watcher,clipboardReader:async()=>'剪贴板里的字'});
  assert.equal((await fetch(base+'/api/pastebar')).status,403);
  assert.equal((await call('/api/pastebar')).data.watching,false);assert.equal(started,0,'默认不监听');
  await call('/api/settings','POST',{patch:{pasteBar:{watch:true}}});assert.equal(started,1);
  emit('会议纪要第一段');emit('Abcd1234!');await new Promise(r=>setTimeout(r,30));
  let st=(await call('/api/pastebar')).data;assert.equal(st.watching,true);assert.deepEqual(st.items.map(x=>x.block.text),['会议纪要第一段']);assert.equal(st.skipped,1);assert.equal(st.items[0].from,'clipboard');
  await call('/api/settings','POST',{patch:{pasteBar:{watch:false}}});assert.equal(stopped,1);
  st=(await call('/api/pastebar/clipboard','POST')).data;assert.equal(st.items[0].block.text,'剪贴板里的字');
  st=(await call('/api/pastebar','POST',{block:{type:'table',rows:[['a','b']]},sourceHost:'et'})).data;assert.equal(st.item.from,'manual');
  st=(await call('/api/pastebar/pin','POST',{id:st.items.at(-1).id,pinned:true})).data;assert.ok(st.items[0].pinned);assert.equal(st.items[0].block.text,'会议纪要第一段');
  st=(await call('/api/pastebar','DELETE')).data;assert.equal(st.items.length,1);
  assert.equal((await call('/api/pastebar?id=nope','DELETE')).status,404);
});
test('剪贴板监听（Windows）：PowerShell 每行输出一段 Base64 文字，服务按行解出；停止时结束进程',()=>{
  const child=new EventEmitter();child.stdout=new EventEmitter();child.stdout.setEncoding=()=>{};let killed=false,args=null;child.kill=()=>{killed=true;};
  const got=[];const w=watchClipboard(t=>got.push(t),{platform:'win32',pid:42,spawnImpl:(cmd,a,o)=>{args={cmd,a,o};return child;}});
  child.stdout.emit('data','ready\n'+Buffer.from('第一条').toString('base64')+'\n'+Buffer.from('多行\r\n文字').toString('base64').slice(0,4));
  child.stdout.emit('data',Buffer.from('多行\r\n文字').toString('base64').slice(4)+'\n');
  assert.deepEqual(got,['第一条','多行\r\n文字']);assert.equal(args.cmd,'powershell.exe');assert.equal(args.o.env.OFFICE_TOOLBOX_PID,'42');assert.ok(args.o.windowsHide);
  assert.match(args.a.at(-1),/GetClipboardSequenceNumber/);assert.ok(!args.a.includes('-EncodedCommand'));
  w.stop();assert.ok(killed);
});

function bridge(a,host){const window={Application:a,confirm:()=>true};vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/'+host+'/',search:''},URLSearchParams,Date,setTimeout,clearTimeout});return window.OfficeTools;}
test('从粘贴栏插入（文字）：插在光标处、不弹确认，格式跟随周围；选中文字时插在选区后面',async()=>{
  const doc={text:'甲乙丙'};let sel=[1,1];
  const range=(s,e)=>{const r={Start:s,End:e,get Text(){return doc.text.slice(r.Start,r.End);},set Text(v){doc.text=doc.text.slice(0,r.Start)+v+doc.text.slice(r.End);r.End=r.Start+v.length;},Select(){sel=[r.Start,r.End];},get Duplicate(){return range(r.Start,r.End);},Collapse(d){if(d===0)r.Start=r.End;else r.End=r.Start;}};return r;};
  const a={ActiveDocument:{Name:'a.docx',ReadOnly:false,Range:range},Selection:{get Range(){return range(sel[0],sel[1]);},Collapse(){sel=[sel[1],sel[1]];}}};
  const T=bridge(a,'wps');let msg=await T.runFromUi('pasteBarInsert',{block:{type:'text',text:'X\nY'}});
  assert.equal(doc.text,'甲X\rY乙丙');assert.match(msg,/格式跟随光标处/);assert.deepEqual(sel,[4,4],'光标在插入内容之后');
  sel=[0,1];await T.runFromUi('pasteBarInsert',{block:{type:'text',text:'Z'}});assert.equal(doc.text,'甲ZX\rY乙丙','选中的文字不被替换');
});
test('从粘贴栏插入（演示）：光标在文本框里时插进文字',async()=>{
  const tr={Text:''};const a={ActivePresentation:{Name:'p.pptx',ReadOnly:false},ActiveWindow:{Selection:{Type:3,TextRange:tr}}};
  const msg=await bridge(a,'wpp').runFromUi('pasteBarInsert',{block:{type:'text',text:'要点'}});assert.equal(tr.Text,'要点');assert.match(msg,/文本框/);
});
test('粘贴栏界面：侧栏加载了脚本；功能区有“粘贴栏”和“放进粘贴栏”，收集不弹确认',()=>{
  const html=fs.readFileSync(new URL('../public/tools.html',import.meta.url),'utf8'),xml=fs.readFileSync(new URL('../public/ribbon.xml',import.meta.url),'utf8'),core=fs.readFileSync(new URL('../public/tools/core.js',import.meta.url),'utf8'),main=fs.readFileSync(new URL('../public/main.js',import.meta.url),'utf8');
  assert.match(html,/<script src="paste-bar\.js"><\/script>/);assert.match(xml,/id="pasteBarOpen"/);assert.match(xml,/id="pasteBarCollect"/);
  assert.match(core,/\['capture','pasteBarCollect'/);assert.match(main,/id==='pasteBarOpen'\)\{open\('files'\)/);
});
