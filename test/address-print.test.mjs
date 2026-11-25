import test from 'node:test';import assert from 'node:assert/strict';import JSZip from 'jszip';
import {parseAddress,withHonorific,verticalNumbers,PAPERS} from '../lib/address-print.mjs';
import {createPrintDocument} from '../lib/print-tools.mjs';

test('地址拆分：空格或 Tab 分隔，邮编位置任意，全角数字也认',()=>{
  assert.deepEqual(parseAddress('〒100-0001 東京都千代田区1-2-3 山田太郎'),{zip:'1000001',address:'東京都千代田区1-2-3',name:'山田太郎'});
  assert.deepEqual(parseAddress('大阪府テスト市2-10\t１２３－４５６７\t株式会社サンプル 御中'),{zip:'1234567',address:'大阪府テスト市2-10',name:'株式会社サンプル 御中'});
  assert.deepEqual(parseAddress('北京市示例路12号 张三 100001','zh'),{zip:'100001',address:'北京市示例路12号',name:'张三'});
  assert.deepEqual(parseAddress('邮编：200000\t上海市示例区88号\t李四','zh'),{zip:'200000',address:'上海市示例区88号',name:'李四'});
  assert.equal(parseAddress('  '),null);
  assert.equal(parseAddress('東京都千代田区').name,'');
});
test('敬称只在没有时补上；竖排数字和字母转成不会横倒的字',()=>{
  assert.equal(withHonorific('山田太郎','ja'),'山田太郎　様');assert.equal(withHonorific('株式会社サンプル 御中','ja'),'株式会社サンプル 御中');
  assert.equal(withHonorific('张三','zh'),'张三 收');assert.equal(withHonorific('李四 亲启','zh'),'李四 亲启');assert.equal(withHonorific('','ja'),'');
  assert.equal(verticalNumbers('1-2-30 5F'),'一ー二ー三〇 五Ｆ');
});
test('はがき：每人一页、纸张 100×148 mm、地址竖排、邮编数字放在规定的 7 个格子里',async()=>{
  const d=await createPrintDocument({kind:'envelope',lang:'zh',paper:'postcard',addresses:'〒100-0001 東京都千代田区1-2-3 山田太郎\n123-4567 大阪府 佐藤',from:'サンプル商事'});
  const xml=await (await JSZip.loadAsync(d.buffer)).file('word/document.xml').async('string');
  assert.equal((xml.match(/<w:pgSz /g)||[]).length,2);assert.match(xml,/w:w="5669"/);assert.match(xml,/w:h="8391"/);
  assert.match(xml,/<w:textDirection w:val="tbRl"\/>/);assert.match(xml,/山田太郎　様/);assert.match(xml,/東京都千代田区一ー二ー三/);
  assert.match(xml,/w:tblpX="2494"/);// 44 mm
  for(const n of '1000001')assert.ok(xml.includes('<w:t xml:space="preserve">'+n+'</w:t>')||xml.includes('<w:t>'+n+'</w:t>'));
  assert.doesNotMatch(xml,/C8102E/);// 默认不画格子
});
test('中国 5 号信封：横排、6 位邮编格子可选、敬称“收”；格子位置表完整',async()=>{
  const d=await createPrintDocument({kind:'envelope',lang:'ja',paper:'cn5',addresses:'北京市示例路12号 张三 100001',from:'示例公司',zipBoxes:true});
  const xml=await (await JSZip.loadAsync(d.buffer)).file('word/document.xml').async('string');
  assert.match(xml,/<w:pgSz w:w="12472" w:h="6236" w:orient="landscape"\/>/);assert.match(xml,/张三 收/);assert.match(xml,/C8102E/);assert.doesNotMatch(xml,/tbRl/);
  for(const p of Object.values(PAPERS))assert.equal(p.zip.x.length,p.style==='zh'?6:7);
  await assert.rejects(createPrintDocument({kind:'envelope',lang:'ja',paper:'postcard',addresses:'\n \n'}),/请填写地址/);
});
test('地址标签：邮编、地址、姓名分行，姓名带敬称',async()=>{
  const d=await createPrintDocument({kind:'labels',lang:'ja',paper:'a4',addresses:'〒100-0001 東京都千代田区1-2-3 山田太郎'});
  const xml=await (await JSZip.loadAsync(d.buffer)).file('word/document.xml').async('string');
  assert.match(xml,/〒100-0001/);assert.match(xml,/山田太郎　様/);
});
