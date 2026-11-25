import test from 'node:test';import assert from 'node:assert/strict';import vm from 'node:vm';import fs from 'node:fs';
const code=fs.readFileSync(new URL('../public/office/taskpane.js',import.meta.url),'utf8');
function setup(){const elements=new Map(),runs=[];let writes=0;const get=id=>{if(!elements.has(id))elements.set(id,{id,hidden:true,textContent:'',className:'',onclick:null});return elements.get(id);};
 const root={document:{getElementById:get,querySelectorAll:()=>[]},Office:{onReady(){}},OfficeAddinTools:{list:()=>[{id:'trimCells',label:'清理空格'},{id:'fillBlanks',label:'空白填上一格'}],run:async(h,id,api,o)=>{runs.push({id,...o});if(id==='fillBlanks'&&!o.confirmed)return {confirm:'将填充测试单元格，是否继续？'};writes++;return '完成';}}};root.window=root;vm.runInNewContext(code,root);return {root,get,runs,writes:()=>writes};}
const tick=()=>new Promise(r=>setImmediate(r));
test('Office 普通工具：显示确认前不执行；取消不修改；接受后执行',async()=>{const s=setup();await s.root.OfficeToolboxPane.go('trimCells',{});assert.equal(s.runs.length,0);assert.equal(s.get('confirmBox').hidden,false);s.get('confirmNo').onclick();assert.equal(s.writes(),0);await s.root.OfficeToolboxPane.go('trimCells',{});s.get('confirmYes').onclick();await tick();assert.equal(s.writes(),1);});
test('Office 通用确认不会跳过工具自身的数据风险确认',async()=>{const s=setup();await s.root.OfficeToolboxPane.go('fillBlanks',{});s.get('confirmYes').onclick();await tick();assert.equal(s.writes(),0);assert.equal(s.runs[0].confirmed,undefined);assert.match(s.get('confirmText').textContent,/填充测试/);s.get('confirmYes').onclick();await tick();assert.equal(s.writes(),1);});

test('Word 与 PowerPoint 工具先显示填写项，再按用户选择执行',async()=>{
  for(const host of ['Word','PowerPoint']) {
    const elements=new Map(),runs=[];
    const make=tag=>({tag,children:[],value:'',hidden:false,textContent:'',append(...xs){this.children.push(...xs);},setAttribute(){}});
    const get=id=>{if(!elements.has(id))elements.set(id,make(id));return elements.get(id);};
    const root={document:{documentElement:{},getElementById:get,querySelectorAll:()=>[],createElement:make},Office:{onReady(fn){this.ready=fn;},context:{displayLanguage:'ja'}},OfficeAddinTools:{list:()=>[{id:'proofCheck',label:'校对',fields:[{key:'mode',label:'模式',value:'comment',options:[{value:'comment',label:'批注'},{value:'fix',label:'改正'}]}]}],batch2:{tools:{word:{proofCheck:1},powerpoint:{proofCheck:1}}},run:async(h,id,api,o)=>{runs.push({h,id,...o});return '完成';}}};
    root.window=root;vm.runInNewContext(code,root);root.Office.ready({host});
    const item=get('tools').children[0],button=item.children[1];button.onclick();assert.equal(runs.length,0);
    const input=item.children[0].children[2].children[0].children[0];input.value='fix';button.onclick();await tick();assert.equal(runs.length,1);assert.equal(runs[0].mode,'fix');
  }
});
