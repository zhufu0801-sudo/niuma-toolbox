// 扫描件版面还原（2026-11-20 体验测试：表格被按列竖排、同一行的栏目被拆散）。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs/promises';import os from 'node:os';import path from 'node:path';import JSZip from 'jszip';
import {layoutOcr} from '../lib/ocr-layout.mjs';
import {createDocTools,layoutText} from '../lib/doc-tools.mjs';
const boxes=JSON.parse(await fs.readFile(new URL('./fixtures/ocr-notice-boxes.json',import.meta.url),'utf8'));
// 模拟 Windows 文字识别的输出顺序：表格区域按列从上到下，栏目行的两半分开出现。
function windowsOrder(lines){const inTable=l=>l.words[0].y>=700&&l.words[0].y<1000,tbl=lines.filter(inTable).sort((a,b)=>a.words[0].x-b.words[0].x||a.words[0].y-b.words[0].y),rest=lines.filter(l=>!inTable(l));return [...rest.slice(0,5),rest[6],rest[5],rest[7],...tbl,...rest.slice(8)];}

test('版面还原：按行排顺序，表格按格子还原，标题、落款、填写栏单独成段',()=>{
  const segs=layoutOcr(windowsOrder(boxes.lines),boxes.width);
  assert.deepEqual(segs[0],{para:{text:'关于年终办公用品申领的通知',level:1}});
  assert.ok(segs.some(s=>s.para&&s.para.text==='申领人：＿＿＿＿＿＿　　部门：＿＿＿＿＿＿'),'同一行的两个栏目放在一起');
  assert.ok(segs.some(s=>s.para&&/^联系电话：＿+　　日期：2026年11月20日$/.test(s.para.text)));
  const table=segs.find(s=>s.table).table;
  assert.deepEqual(table,[['序号','物品名称','规格','数量','备注'],['1','A4复印纸','70g','10箱',''],['2','中性笔','0.5mm黑色','50支',''],['3','文件夹','A4蓝色','30个','加厚']]);
  assert.deepEqual(segs.slice(-2),[{para:{text:'示例公司行政部',align:'right'}},{para:{text:'2026年11月20日',align:'right'}}]);
  assert.ok(!segs.some(s=>s.para&&/各部门：＿/.test(s.para.text)),'单独一行的称呼不补下划线');
  assert.match(layoutText(segs),/序号\t物品名称\t规格\t数量\t备注\n1\tA4复印纸/);
});
test('扫描件 PDF 转 Word：用 Windows 识别的词位置写出真正的 Word 表格和居中标题',async t=>{
  const d=await fs.mkdtemp(path.join(os.tmpdir(),'scan-'));t.after(()=>fs.rm(d,{recursive:true,force:true}));
  const pdf=path.join(d,'扫描通知.pdf');await fs.copyFile(new URL('./fixtures/scan-notice.pdf',import.meta.url),pdf);
  const exec=async(cmd,args,o)=>({stdout:JSON.stringify({lang:'zh-Hans-CN',langs:['zh-Hans-CN'],results:o.env.OCR_FILES.split('|').map(f=>({file:f,width:boxes.width,lines:windowsOrder(boxes.lines)}))})});
  const r=await createDocTools({platform:'win32',exec,tmpRoot:d}).pdfToWord({path:pdf,lang:'zh'});
  assert.equal(r.ocr,1);assert.equal(r.tables,1);
  const xml=await (await JSZip.loadAsync(await fs.readFile(r.path))).file('word/document.xml').async('string');
  assert.match(xml,/<w:tbl>/);assert.match(xml,/<w:jc w:val="center"\/>[\s\S]*关于年终办公用品申领的通知/);assert.match(xml,/<w:jc w:val="right"\/>[\s\S]*示例公司行政部/);
  const order=['申领人','部门：＿','联系电话','序号','A4复印纸','特此通知'].map(s=>xml.indexOf(s));assert.deepEqual(order,[...order].sort((a,b)=>a-b),'阅读顺序和原件一致');
});
