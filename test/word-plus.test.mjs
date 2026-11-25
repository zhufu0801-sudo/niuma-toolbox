import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
import {makeApp,makeBook} from './helpers/et-mock.mjs';
function load(a){const asked=[];const window={Application:a,confirm:m=>{asked.push(m);return true;}};vm.runInNewContext(toolsCode(),{window,location:{pathname:'/addon/wps/',search:''},URLSearchParams,Date});return {T:window.OfficeTools,W:window.OfficeWordPlus,asked};}
test('上下标：单位改上标、化学式改下标，B2B、A4、MP3、1.5m2 之外的编号不动',()=>{
  const {W}=load({});const m=W.marks('面积 120m2，体积 3cm3；水是 H2O，CO2 和 H2SO4；O2 充足。B2B 业务，A4 纸，MP3，Fe2O3。');
  const pick=x=>m.filter(k=>k.sup===x).map(k=>k.word);
  assert.deepEqual([...new Set(pick(true))],['m2','cm3']);assert.deepEqual([...new Set(pick(false))],['H2O','CO2','H2SO4','O2','Fe2O3']);
  assert.equal(m.filter(k=>k.word==='H2SO4').length,2);assert.equal(W.chemDigits('B2B'),null);assert.equal(W.chemDigits('COVID19'),null);assert.deepEqual([...W.chemDigits('N2')],[1]);assert.equal(W.chemDigits('C2'),null);
});
test('上下标：按位置设置字体，核对文字后才改',async()=>{
  const text='H2O 和 5m2',fonts={};const d={Content:{Start:0,End:text.length,get Text(){return text;}},Range:(s,e)=>({get Text(){return text.slice(s,e);},Font:fonts[s]||(fonts[s]={})})};
  const a={ActiveDocument:d,Selection:{Text:text,Range:{Start:0,End:text.length,Text:text}}};const {T}=load(a);
  const r=await T.run('superSubscript',{host:'wps'});assert.match(r,/已设置 2 处上标 \/ 下标：m2、H2O|已设置 2 处上标 \/ 下标：H2O、m2/);assert.equal(fonts[1].Subscript,true);assert.equal(fonts[8].Superscript,true);assert.equal(fonts[6],undefined);
});
test('输入后自动加边框：公式以选区左上角为准，带四条边框',async()=>{
  const book=makeBook({S:[['a','b'],['','']]}),app=makeApp(book);app.select('A2:B300');const {T}=load(app);const r=await T.run('autoBorder',{host:'et'});
  assert.match(r,/A2:B300/);const cf=book.list[0].cf.at(-1);assert.equal(cf.formula,'=COUNTA($A2:$B2)>0');
});
// 简单的文字文档模拟：文字存在一个字符串里，段落按 \r 分开；heads 是标题段落、centers 是居中段落（按开头文字认）。
function textDoc(text,{heads=[],centers=[]}={}){
  const doc={text,fmts:{}};const paraAt=s=>{const parts=doc.text.split('\r');let p=0;for(let i=0;i<parts.length;i++){const e=p+parts[i].length+1;if(s<e)return {i,start:p,end:e,text:parts[i]};p=e;}return null;};
  const range=(s,e)=>({get Start(){return s;},get End(){return e;},get Text(){return doc.text.slice(s,e);},set Text(v){doc.text=doc.text.slice(0,s)+v+doc.text.slice(e);e=s+v.length;},Delete(){doc.text=doc.text.slice(0,s)+doc.text.slice(e);e=s;},
    get ParagraphFormat(){const p=paraAt(s),k=p.text.replace(/^[ 　\t]+/,'');return doc.fmts[k]||(doc.fmts[k]={Alignment:centers.some(c=>p.text.trim().startsWith(c))?1:0,TabStops:{list:[],ClearAll(){this.list=[];},Add(x){this.list.push(x);}}});},Font:{Size:12},Information:()=>false,ListFormat:{ListType:0},
    get Paragraphs(){const list=[];let p=0;doc.text.split('\r').forEach((t,i,arr)=>{const st=p,en=p+t.length+(i<arr.length-1?1:0);p=st+t.length+1;if(en>s&&st<e&&!(i===arr.length-1&&!t))list.push({OutlineLevel:heads.some(h=>t.trim().startsWith(h))?1:10,Range:range(st,en)});});return {Count:list.length,Item:i=>list[i-1]};}});
  doc.Range=range;Object.defineProperty(doc,'Content',{get:()=>range(0,doc.text.length)});doc.PageSetup={PageWidth:595,LeftMargin:90,RightMargin:90};return doc;
}
test('段落规范化：开头空格改首行缩进、Shift+Enter 改分段，标题和居中行不动',async()=>{
  const d=textDoc('一、概况\r　　第一段正文。\x0b第二句另起一行。\r    第三段\r落款单位\r',{heads:['一、'],centers:['落款']});const a={ActiveDocument:d,Selection:{Text:''}};const {T}=load(a);
  const r=await T.run('paraNormalize',{host:'wps'});assert.match(r,/已整理 3 个正文段落/);assert.match(r,/删掉了 2 段开头/);assert.match(r,/1 处手动换行符/);
  assert.equal(d.text,'一、概况\r第一段正文。\r第二句另起一行。\r第三段\r落款单位\r');assert.deepEqual(['第一段正文。','第二句另起一行。','第三段'].map(k=>d.fmts[k].CharacterUnitFirstLineIndent),[2,2,2]);assert.equal(d.fmts['一、概况']&&d.fmts['一、概况'].CharacterUnitFirstLineIndent,undefined);assert.equal(d.fmts['落款单位'].CharacterUnitFirstLineIndent,undefined);
});
test('删除括号里的答案：只清空选项和对错，其他括号内容不动',async()=>{
  const d=textDoc('1. 下列正确的是（ B ）。\r2. 地球是圆的。（√）\r3. 多选（ACD）\r4. 备注（见附件）（2024年）\r');const a={ActiveDocument:d,Selection:{Text:''}};const {T,asked}=load(a);
  assert.match(await T.run('deleteAnswers',{host:'wps'}),/已清空 3 处答案/);assert.match(asked.at(-1),/将清空 3 处/);
  assert.equal(d.text,'1. 下列正确的是（　　）。\r2. 地球是圆的。（　　）\r3. 多选（　　）\r4. 备注（见附件）（2024年）\r');
});
test('合同落款：甲乙方之间的空格换成制表位',async()=>{
  const d=textDoc('甲方：某某公司        乙方：某某商店\r日期：          日期：\r');const a={ActiveDocument:d};a.Selection={get Text(){return d.text;},get Range(){return d.Content;}};const {T}=load(a);
  await T.run('contractSign',{host:'wps'});assert.equal(d.text,'甲方：某某公司\t乙方：某某商店\r日期：\t日期：\r');assert.deepEqual(d.fmts['甲方：某某公司\t乙方：某某商店'].TabStops.list,[208]);
});
