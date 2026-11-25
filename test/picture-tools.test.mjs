import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {toolsCode} from './helpers/tools-code.mjs';
const code=toolsCode();
function run(app,id,host){const window={Application:app,confirm:()=>true};vm.runInNewContext(code,{window,fetch:async()=>{throw Error('no network');},location:{pathname:'/addon/'+host+'/'},setTimeout,clearTimeout});return window.OfficeTools.run(id,{host});}
function word(pics){const list=pics.map(([w,h])=>({Width:w,Height:h}));const range={Start:0,End:10,InlineShapes:{Count:list.length,Item:i=>list[i-1]}};
 const ps={PageWidth:595,LeftMargin:90,RightMargin:90};const doc={Name:'a.docx',FullName:'C:/a.docx',PageSetup:ps};
 return {list,app:{ActiveDocument:doc,Selection:{Text:'\u0001',Range:range,Sections:{Item:()=>({PageSetup:ps})}}}};}
test('图片统一大小：以第一张宽度为准，按比例改高度',async()=>{
 const w=word([[200,100],[400,400],[100,50]]);assert.match(await run(w.app,'unifyPictures','wps'),/3 张/);
 assert.deepEqual(w.list.map(p=>[p.Width,p.Height]),[[200,100],[200,200],[200,100]]);
 await assert.rejects(run(word([[1,1]]).app,'unifyPictures','wps'),/至少 2 张/);
 await assert.rejects(run(word([]).app,'unifyPictures','wps'),/没有嵌入型图片/);
});
test('图片不超出页面：只缩小超出版心的图片',async()=>{
 const w=word([[600,300],[300,100]]);assert.match(await run(w.app,'fitPictures','wps'),/1 张/);
 assert.deepEqual(w.list.map(p=>[Math.round(p.Width),Math.round(p.Height)]),[[415,208],[300,100]]);
});
test('证件照按实际尺寸：比例相符才改，最多 2 张',async()=>{
 const w=word([[317,200]]);assert.match(await run(w.app,'idCardSize','wps'),/8.56×5.4/);assert.equal(Math.round(w.list[0].Width*100)/100,242.65);
 await assert.rejects(run(word([[300,300]]).app,'idCardSize','wps'),/长宽比/);
 await assert.rejects(run(word([[317,200],[317,200],[317,200]]).app,'idCardSize','wps'),/最多处理 2 张/);
});
test('图片适应单元格：只处理左上角在选区内的图片，缩放居中并随单元格移动',async()=>{
 const cell={Row:2,Column:2,Left:100,Top:50,Width:80,Height:40};cell.MergeArea=cell;
 const inside={Type:13,Width:160,Height:40,TopLeftCell:cell},outside={Type:13,Width:50,Height:50,TopLeftCell:{Row:9,Column:9}};
 const ws={Name:'S',Shapes:{Count:2,Item:i=>[inside,outside][i-1]}};
 const range={Row:1,Column:1,Rows:{Count:3},Columns:{Count:3},Address:()=>'$A$1:$C$3'};
 const app={ActiveWorkbook:{Name:'a.xlsx',FullName:'C:/a.xlsx'},ActiveSheet:ws,Selection:range};
 assert.match(await run(app,'fitPicturesToCells','et'),/1 张/);
 assert.equal(inside.Width,76);assert.equal(inside.Height,19);assert.equal(inside.Left,102);assert.equal(inside.Top,60.5);assert.equal(inside.Placement,1);assert.equal(outside.Width,50);
});
