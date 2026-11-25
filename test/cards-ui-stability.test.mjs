import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {TOOL_MODULES} from './helpers/tools-code.mjs';
import {createClock,createDocument,createSharedBrowser,loadScripts,flush} from './helpers/frontend-fakes.mjs';

test('卡片列表内容不变时不替换节点，定时刷新不打断点击或拖拽；无功能区按钮的动作不列入开关',async()=>{
 const clock=createClock(),document=createDocument(),browser=createSharedBrowser(),writes=[];
 const a={ActiveDocument:{Name:'独立测试.docx'},Selection:{Range:{Start:1,End:1,Duplicate:{Collapse(){},set Text(value){writes.push(value);}}}}};
 const card={id:'one',title:'测试卡片',created:1,sourceHost:'wps',block:{type:'text',text:'安全插入'}};
 const root=loadScripts({confirm:()=>true,...clock,...browser,URLSearchParams,document,Application:a,location:{origin:'http://127.0.0.1:39127',pathname:'/tools.html',search:'?host=wps&view=cards'},history:{replaceState(){}},ToolIcons:{path:id=>id},PaneBridge:{listen(){}},StudioLaunch:{open(){}},fetch:async url=>({ok:true,json:async()=>String(url).includes('/session')?{token:'t'}:{cards:[card]}})},[...TOOL_MODULES,'tools-ui.js']);
 await flush();const list=document.getElementById('cardList'),first=list.children[0],created=document.created.length;assert.ok(first);
 await clock.advance(8000);assert.equal(list.children.length,1);assert.equal(list.children[0],first);assert.equal(document.created.length,created);
 first.ondragstart({dataTransfer:{setData(){}}});await clock.advance(4000);assert.equal(list.children[0],first);first.ondragend();
 first.children[2].children[0].click();await flush();assert.deepEqual(writes,['安全插入']);
 const labels=document.getElementById('manageList').children.map(label=>label.children[1].textContent);assert.ok(!labels.includes('插入内容卡片'));assert.ok(labels.includes('设为内容模块'));assert.equal(root.OfficeTools.version,JSON.parse(fs.readFileSync(new URL('../package.json',import.meta.url),'utf8')).version);
});
