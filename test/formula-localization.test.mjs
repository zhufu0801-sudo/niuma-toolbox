import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
test('formula defaults follow UI language until the user edits them',()=>{
 const nodes=new Map();let lang='ja',changed;
 function node(){return {value:'',hidden:false,children:[],dataset:{},listeners:{},appendChild(n){this.children.push(n);if(!this.value)this.value=n.value;},addEventListener(k,fn){(this.listeners[k]??=[]).push(fn);},querySelectorAll(){return [];},parentNode:{insertBefore(){}}};}
 const document={createElement:node,getElementById(id){if(!nodes.has(id))nodes.set(id,node());return nodes.get(id);}};
 const c={document,localStorage:{getItem:()=>null},OfficeTools:{host:()=> 'et'},OfficeFormula:{types:{lookup:'查找'}},OfficeFormulaHost:{},OfficeI18n:{lang:()=>lang,onChange:fn=>changed=fn}};c.window=c;
 vm.runInNewContext(fs.readFileSync(new URL('../public/formula-panel.js',import.meta.url),'utf8'),c);
 assert.equal(nodes.get('faFallback').value,'見つかりません');assert.equal(nodes.get('faYes').value,'はい');
 nodes.get('faFallback').value='客户自定义';for(const fn of nodes.get('faFallback').listeners.input)fn();
 lang='zh';changed();assert.equal(nodes.get('faFallback').value,'客户自定义');assert.equal(nodes.get('faYes').value,'是');assert.equal(nodes.get('faNo').value,'否');
 lang='ja';changed();assert.equal(nodes.get('faFallback').value,'客户自定义');assert.equal(nodes.get('faNo').value,'いいえ');
});
