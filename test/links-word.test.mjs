import test from 'node:test';import assert from 'node:assert/strict';import vm from 'node:vm';import {createCoachServer} from '../server.mjs';import {toolsCode} from './helpers/tools-code.mjs';
function document(name,text=''){
 const marks=new Map(),font={Name:'宋体',NameFarEast:'宋体',Size:12,Bold:0,Italic:0,Color:0},pf={Alignment:0};
 const d={Name:name,Path:'C:/docs',FullName:'C:/docs/'+name,ReadOnly:false,TrackRevisions:false,text,font,marks,writes:0,failNext:false};
 d.Range=(start,end)=>{const range={Start:start,End:end,StoryType:1,Font:font,ParagraphFormat:pf,Fields:{Count:0},Tables:{Count:0},Collapse(){this.Start=this.End;}};
  Object.defineProperty(range,'Text',{get(){return d.text.slice(range.Start,range.End);},set(value){d.writes++;const oldEnd=range.End,delta=value.length-(range.End-range.Start);d.text=d.text.slice(0,range.Start)+value+d.text.slice(range.End);range.End=range.Start+value.length;for(const item of marks.values()){if(item.start>=oldEnd){item.start+=delta;item.end+=delta;}else if(item.start<=range.Start&&item.end>=oldEnd)item.end+=delta;}if(d.failNext){d.failNext=false;throw Error('模拟写入中断');}}});
  Object.defineProperty(range,'Duplicate',{get:()=>d.Range(range.Start,range.End)});range.Paragraphs={Item:()=>({OutlineLevel:10,Range:range})};return range;};
 d.Bookmarks={Exists:name=>marks.has(name),Add(name,range){marks.set(name,{start:range.Start,end:range.End});return this.Item(name);},Item(name){if(!marks.has(name))throw Error('书签不存在');return {get Range(){const item=marks.get(name);return d.Range(item.start,item.end);},Delete(){marks.delete(name);}};}};return d;
}
function selection(d,start,end){const range=d.Range(start,end);return {Range:range,Tables:{Count:0},get Text(){return range.Text;}};}
test('Word linked workflow captures source, inserts at target, updates only the bookmark and protects local edits',async()=>{
 const server=createCoachServer();await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));try{
 const base='http://127.0.0.1:'+server.address().port,source=document('source.docx','原文'),target=document('target.docx','前后'),docs=[source,target];const app={ActiveDocument:source,Selection:selection(source,0,2),Documents:{Count:2,Item:i=>docs[i-1]}};
 const root={Application:app,confirm:()=>true};vm.runInNewContext(toolsCode(),{window:root,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date,setTimeout,clearTimeout,fetch:(url,options)=>fetch(base+url,options)});const T=root.OfficeTools;
 assert.match(await T.run('saveLinkedModule'),/已设为联动模块/);let card=(await T.cards.list())[0];assert.equal(card.source.path,source.FullName);assert.ok(source.marks.has(card.source.mark));
 app.ActiveDocument=target;app.Selection=selection(target,1,1);assert.match(await T.run('insertLinkedModule',{cardId:card.id}),/已联动插入/);assert.equal(target.text,'前原文后');card=(await T.cards.list())[0];assert.equal(card.targets.length,1);
 source.text='新内容';source.marks.get(card.source.mark).end=3;assert.match(await T.run('updateLinkedModules'),/已更新/);assert.equal(target.text,'前新内容后');card=(await T.cards.list())[0];assert.equal(card.targets[0].moduleRevision,2);
 const linkedRange=target.Bookmarks.Item(card.targets[0].anchor.mark).Range;linkedRange.Text='手动修改';const before=target.text,writes=target.writes;source.text='第三版';source.marks.get(card.source.mark).end=3;assert.match(await T.run('updateLinkedModules'),/已被修改/);assert.equal(target.text,before);assert.equal(target.writes,writes);
 }finally{await new Promise(resolve=>server.close(resolve));}
});
test('Word link update cancellation, lost anchors and native write failures do not silently replace content',async()=>{
 const d=document('target.docx','旧内容'),app={ActiveDocument:d};d.Bookmarks.Add('OTD_test',d.Range(0,3));let confirmed=true,acks=0;const root={Application:app,confirm:()=>confirmed};vm.runInNewContext(toolsCode(),{window:root,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date,setTimeout,clearTimeout,fetch:async url=>{if(url==='/api/session')return Response.json({token:'t'});acks++;return Response.json({link:{}});}});const T=root.OfficeTools,anchor={host:'wps',path:d.FullName,mark:'OTD_test'},baseline=T.links.readTarget(d,anchor).snapshot,card={id:'module',title:'测试',revision:2,source:{host:'et'},block:{type:'text',text:'新的内容'}},link={id:'abc',anchor,baseline,moduleRevision:1,revision:1},c={a:app,d,h:'wps'};
 confirmed=false;assert.equal(await T.links.updateOne(c,card,link,true),'已取消');assert.equal(d.text,'旧内容');assert.equal(d.writes,0);
 confirmed=true;d.failNext=true;await assert.rejects(T.links.updateOne(c,card,link,true),/已尝试恢复/);assert.equal(d.text,'旧内容');assert.equal(acks,0);
 d.marks.delete('OTD_test');await assert.rejects(T.links.updateOne(c,card,link,true),/书签已删除/);assert.equal(d.text,'旧内容');
});
test('cross-host update waits for a fresh source check, and closed sources never overwrite targets',async()=>{
 const d=document('target.docx','旧内容'),app={ActiveDocument:d};d.Bookmarks.Add('OTD_test',d.Range(0,3));let time=1000,allow=false,checks=0;
 class ClockDate extends Date{static now(){return time;}}
 const root={Application:app,confirm:()=>true};vm.runInNewContext(toolsCode(),{window:root,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date:ClockDate,setTimeout:fn=>{time+=500;queueMicrotask(fn);},clearTimeout(){},fetch:async url=>url==='/api/session'?Response.json({token:'t'}):Response.json({link:{}})});const T=root.OfficeTools,anchor={host:'wps',path:d.FullName,mark:'OTD_test'},link={id:'link',anchor,baseline:T.links.readTarget(d,anchor).snapshot,moduleRevision:1,revision:1},card={id:'module',title:'跨宿主模块',source:{host:'et'},revision:2,block:{type:'text',text:'新内容'},targets:[link]};T.cards.list=async()=>{checks++;return [{...card,sourceCheckedAt:allow&&checks>2?time+1:1}];};
 assert.match(await T.run('updateLinkedModules'),/无法读取来源最新内容/);assert.equal(d.writes,0);assert.equal(d.text,'旧内容');
 allow=true;checks=0;assert.match(await T.run('updateLinkedModules'),/已更新/);assert.equal(d.text,'新内容');assert.ok(checks>2);
});
