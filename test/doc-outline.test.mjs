import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const ctx={};ctx.globalThis=ctx;vm.runInNewContext(fs.readFileSync(new URL('../public/doc-outline.js',import.meta.url),'utf8'),ctx);const O=ctx.OfficeDocOutline;
const P=(text,level=10,style='正文')=>({text,level,style});
test('按标题样式整理：标题、副标题、每节要点，长段落取第一句',()=>{
  const r=O.build([P('目录 1 概况',10,'目录 1'),P('第三季度工作汇报',10,'标题'),P('销售部 2026 年 10 月'),P('一、总体情况',1,'标题 1'),P('本季度销售额 1200 万元，同比增长 15%。主要来自华东地区的新客户。'),P('重点项目',2,'标题 2'),P('二、存在问题',1,'标题 1'),P('人手不足'),P('三、空标题',1,'标题 1'),P('四、下季度计划',1,'标题 1'),P('扩大华南市场')]);
  assert.equal(r.title,'第三季度工作汇报');assert.equal(r.subtitle,'销售部 2026 年 10 月');
  assert.deepEqual([...r.sections.map(s=>s.heading)],['总体情况','存在问题','下季度计划']);
  assert.deepEqual([...r.sections[0].points],['本季度销售额 1200 万元，同比增长 15%。','重点项目']);assert.equal(r.skipped,1);
  assert.equal(r.outline.split('\n\n')[0],'总体情况\n- 本季度销售额 1200 万元，同比增长 15%。\n- 重点项目');
  assert.match(O.summary(r),/整理出 3 节：“总体情况”、“存在问题”、“下季度计划”。\n1 个标题下面没有内容/);
});
test('没有标题样式：按“一、”“（一）”“1.”编号认标题；没有标题时用文件名',()=>{
  const r=O.build([P('一、背景'),P('（一）市场'),P('需求旺盛。'),P('二、方案'),P('1. 第一步'),P('分三步走。')],{name:'方案.docx'});
  assert.equal(r.numbered,true);assert.equal(r.title,'方案');assert.deepEqual([...r.sections.map(s=>s.heading)],['背景','方案']);
  assert.deepEqual([...r.sections[0].points],['市场','需求旺盛。']);assert.deepEqual([...r.sections[1].points],['第一步','分三步走。']);
  assert.equal(O.numbered('一、这是一个以句号结尾的长句子。'),0);assert.equal(O.numbered('第二章 方法'),1);assert.equal(O.numbered('1.5 倍'),0);
});
test('出错提示和上限：没有标题、标题下都没内容、最多 12 节、每节最多 10 条',()=>{
  assert.throws(()=>O.build([P('只有正文。'),P('还是正文。')]),/没有找到标题/);
  assert.throws(()=>O.build([P('甲',1,'标题 1'),P('乙',1,'标题 1')]),/都没有内容/);
  assert.throws(()=>O.build([P('  ')]),/没有文字/);
  const many=[];for(let i=1;i<=14;i++){many.push(P('第'+i+'节',1,'Heading 1'));for(let k=0;k<12;k++)many.push(P('要点'+i+'-'+k));}
  const r=O.build(many,{name:'x.docx'});assert.equal(r.sections.length,12);assert.equal(r.cut,2);assert.equal(r.sections[0].points.length,10);
  const long=O.build([P('标题',1,'Heading 1'),P('甲'.repeat(80))]);assert.equal(long.sections[0].points[0].length,60);
});
