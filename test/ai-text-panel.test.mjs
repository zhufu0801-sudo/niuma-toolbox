import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const src=fs.readFileSync(new URL('../public/ai-text-panel.js',import.meta.url),'utf8');
function load(host,app){
  const stub=()=>({value:'',textContent:'',hidden:false,disabled:false,readOnly:false,className:'',appendChild(){},onclick:null,onchange:null});
  const els={};const section={className:'',dataset:{},set innerHTML(h){for(const m of h.matchAll(/id="(at[A-Za-z]+)"/g))els[m[1]]=stub();},querySelector:s=>els[s.slice(1)]};
  const w={OfficeTools:{host:()=>host,app:()=>app,lib:{ask:async()=>true}},localStorage:{getItem:()=>null,setItem(){}},fetch:async()=>({json:async()=>({config:{}})}),
    document:{createElement:t=>t==='section'?section:stub(),getElementById:id=>id==='replaceSection'?{parentNode:{insertBefore(){}}}:null,querySelector:()=>null,addEventListener(){}}};
  w.window=w;vm.runInNewContext(src,w);return {host:w.AiTextPanel.wpsHost,els,w};
}
test('WPS 文字：读取选中文字，替换或插在后面',async()=>{
  const range={Text:'',after:'',InsertAfter(x){this.after+=x;}};const app={Selection:{Text:'你好，世界\r',Range:range}};
  const {host,els}=load('wps',app);assert.ok(els.atAction&&els.atOut);const s=host.read();assert.equal(s.text,'你好，世界');
  assert.match(host.write(s,'こんにちは\n世界','replace'),/替换/);assert.equal(range.Text,'こんにちは\r世界');
  host.write(s,'Hello','insert');assert.equal(range.after,'\rHello');
  assert.throws(()=>load('wps',{Selection:{Text:'\r',Range:range}}).host.read(),/选中/);
});
test('WPS 表格：只取文字单元格（跳过公式和数字），结果写到右侧一列，覆盖前先确认',async()=>{
  const grid=[[{Value2:'销售部'},{Value2:1200}],[{Value2:'=A1',HasFormula:true},{Value2:'完成'}]];
  const right=[[{Value2:''},{Value2:''}],[{Value2:'不动'},{Value2:'旧'}]];
  const mk=g=>({Rows:{Count:2},Columns:{Count:2},Cells:{Count:4,Item:(i,j)=>g[i-1][j-1]}});
  const sel={...mk(grid),Address:()=>'A1:B2',Offset:(r,c)=>{assert.equal(c,2);return mk(right);}};
  const {host}=load('et',{Selection:sel});const s=host.read();assert.equal(s.items.map(x=>x.text).join(),'销售部,完成');assert.equal(s.address,'A1:B2');
  const ask=host.write(s,['営業部','=完了'],'insert');assert.match(ask.confirm,/1 个/);
  host.write(s,['営業部','=完了'],'insert',true);assert.equal(right[0][0].Value2,'営業部');assert.equal(right[1][1].Value2,"'=完了");assert.equal(grid[0][0].Value2,'销售部');assert.equal(right[1][0].Value2,'不动');
  host.write(s,['営業部','完了'],'replace');assert.equal(grid[0][0].Value2,'営業部');assert.equal(grid[1][1].Value2,'完了');assert.equal(grid[0][1].Value2,1200);
});
test('WPS 演示：读取选中文本框的文字并写回',async()=>{
  const tr={Text:'季度汇报',InsertAfter(x){this.Text+=x;}};const {host}=load('wpp',{ActiveWindow:{Selection:{TextRange:tr}}});
  const s=host.read();assert.equal(s.text,'季度汇报');host.write(s,'四半期報告','replace');assert.equal(tr.Text,'四半期報告');host.write(s,'Q3','insert');assert.equal(tr.Text,'四半期報告\rQ3');
});
test('写作方式：可以从技巧导航直接选中，续写时不显示“替换”',async()=>{
  const {els,w}=load('wps',{Selection:{Text:'',Range:{}}});
  assert.equal(w.AiTextPanel.select('continue'),true);assert.equal(els.atAction.value,'continue');assert.equal(els.atReplace.hidden,true);assert.equal(els.atExtraBox.hidden,false);assert.equal(els.atInsert.textContent,'接在原文后');
  assert.equal(w.AiTextPanel.select('zh2ja'),true);assert.equal(els.atReplace.hidden,false);assert.equal(els.atExtraBox.hidden,true);
  assert.equal(w.AiTextPanel.select('nope'),false);
});
test('AI 对话“插入文档”：文字插在光标处，表格向下写（单个公式按公式写），演示没选文字就加文本框',async()=>{
  const range={Start:5,End:5,after:'',InsertAfter(x){this.after+=x;}};let {host}=load('wps',{Selection:{Range:range}});
  assert.match(host.insert('第一行\n第二行\n'),/已插入/);assert.equal(range.after,'第一行\r第二行');
  range.End=8;range.after='';host.insert('补充');assert.equal(range.after,'\r补充');
  assert.throws(()=>host.insert('  '),/没有可插入/);
  const cells=[{Value2:''},{Value2:'旧'},{Value2:''}];const cell={Formula:'',Address:()=>'B2',Resize:(r,c)=>{assert.equal(c,1);return {Cells:{Item:i=>cells[i-1]}};}};
  ({host}=load('et',{ActiveCell:cell}));let ask=host.insert('甲\n=乙\n丙');assert.match(ask.confirm,/B2 开始向下写入 3 行，其中 1 个/);
  assert.match(host.insert('甲\n=乙\n丙',true),/写入 3 行/);assert.deepEqual(cells.map(x=>x.Value2),['甲',"'=乙",'丙']);
  ask=host.insert('=SUM(A1:A3)');assert.match(ask.confirm,/公式写入 B2/);host.insert('=SUM(A1:A3)',true);assert.equal(cell.Formula,'=SUM(A1:A3)');
  const shapes=[];const slide={Shapes:{AddTextbox:(o,l,t,w,h)=>{const b={TextFrame:{TextRange:{Text:''}}};shapes.push(b);return b;}}};
  ({host}=load('wpp',{ActiveWindow:{Selection:{get TextRange(){throw Error('no text');}},View:{Slide:slide}}}));assert.match(host.insert('要点一\n要点二'),/文本框/);assert.equal(shapes[0].TextFrame.TextRange.Text,'要点一\r要点二');
  const tr={Text:'标题',InsertAfter(x){this.Text+=x;}};({host}=load('wpp',{ActiveWindow:{Selection:{TextRange:tr}}}));host.insert('副标题');assert.equal(tr.Text,'标题\r副标题');
});
