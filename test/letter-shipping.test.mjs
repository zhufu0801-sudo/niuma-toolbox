import test from 'node:test';import assert from 'node:assert/strict';import JSZip from 'jszip';
import fs from 'node:fs/promises';import os from 'node:os';import path from 'node:path';
import {letterLines,kanjiNumber,wareki,fullWidth} from '../lib/letter-layout.mjs';
import {greetingText,GREETINGS} from '../lib/letter-templates.mjs';
import {parseShipping,shippingSections} from '../lib/address-print.mjs';
import {createPrintDocument} from '../lib/print-tools.mjs';
import {createTemplateLibrary} from '../lib/template-library.mjs';
const xmlOf=async b=>(await JSZip.loadAsync(b)).file('word/document.xml').async('string');

test('和暦と漢数字：位取りで書く（十一月十九日）、元年も対応',()=>{
  assert.equal(kanjiNumber(8),'八');assert.equal(kanjiNumber(11),'十一');assert.equal(kanjiNumber(20),'二十');assert.equal(kanjiNumber(19),'十九');
  assert.equal(wareki(new Date(2026,10,19)),'令和八年十一月十九日');assert.equal(wareki(new Date(2019,4,1),{kichijitsu:true}),'令和元年五月吉日');
  assert.equal(fullWidth('3F 12-3'),'３Ｆ １２－３');
});
test('挨拶状の例文：時候の挨拶と日付は当日から、頭語と結語が揃う',()=>{
  const t=greetingText('moving',new Date(2026,10,19));assert.match(t,/^拝啓　晩秋の候/);assert.match(t,/\n敬具\n令和八年十一月吉日\n/);assert.match(t,/\n記\n/);assert.match(t,/以上$/);
  assert.match(greetingText('thanks',new Date(2026,0,5)),/厳寒の候/);assert.match(greetingText('yearend',new Date(2026,10,19)),/令和八年十二月吉日/);
  for(const g of GREETINGS)assert.ok(greetingText(g.id).length>40,g.id);assert.throws(()=>greetingText('nope'),/例文/);
});
test('縦書きの段落の役割：結語は下、記は中央、結語後の日付・差出人・宛名',()=>{
  const roles=letterLines(greetingText('thanks',new Date(2026,10,19))).map(l=>l.role);
  assert.deepEqual(roles.slice(-4),['closing','date','sender','addressee']);
  const m=letterLines(greetingText('moving',new Date(2026,10,19)));const by=r=>m.filter(l=>l.role===r).map(l=>l.text);
  assert.deepEqual(by('closing'),['敬具','以上']);assert.deepEqual(by('ki'),['記']);assert.deepEqual(by('sender'),['〇〇株式会社','代表取締役　〇〇　〇〇']);
  assert.ok(m.find(l=>l.text.startsWith('新住所')).role==='body');
});
test('縦書き文書：tbRl、結語は下揃え、数字は全角、表題を中央に',async()=>{
  const d=await createPrintDocument({kind:'vertical',lang:'ja',subject:'移転のご挨拶',message:'拝啓　本文です。3階に移転します。\n敬具\n令和八年十一月吉日\n〇〇株式会社'});
  const x=await xmlOf(d.buffer);assert.match(x,/<w:textDirection w:val="tbRl"\/>/);assert.match(x,/３階/);assert.match(x,/<w:jc w:val="end"\/><\/w:pPr><w:r><w:t[^>]*>敬具/);assert.match(x,/<w:jc w:val="center"\/>[\s\S]*移転のご挨拶/);
});
test('配送ラベル：電話番号を見分ける（中国の携帯・固定、日本の番号）',()=>{
  assert.deepEqual(parseShipping('100001 北京市测试区示例路12号 张三 13800000000','zh'),{zip:'100001',address:'北京市测试区示例路12号',name:'张三',tel:'13800000000'});
  assert.equal(parseShipping('邮编：200000\t上海市示例区88号\t李四\t021-12345678','zh').tel,'021-12345678');
  const j=parseShipping('〒100-0001 東京都千代田区1-2-3 山田太郎 TEL 03-1234-5678','ja');assert.deepEqual(j,{zip:'1000001',address:'東京都千代田区1-2-3',name:'山田太郎',tel:'03-1234-5678'});
  assert.equal(parseShipping('東京都千代田区1-2-3 山田太郎','ja').tel,'');assert.equal(parseShipping('  ','ja'),null);
});
test('配送ラベル：A4 は 1 枚 4 件（点線で切る）、100×150 は 1 枚 1 件',async()=>{
  const five=Array.from({length:5},(_,i)=>'10000'+i+' 北京市示例路'+i+'号 张'+i+' 1380000000'+i).join('\n');
  assert.equal(shippingSections({paper:'a4',addresses:five,lang:'zh'}).length,2);assert.equal(shippingSections({paper:'label100',addresses:five,lang:'zh'}).length,5);
  assert.throws(()=>shippingSections({addresses:'\n'}),/请填写地址/);
  const d=await createPrintDocument({kind:'shipping',lang:'ja',paper:'label100',addresses:'〒100-0001 東京都千代田区1-2-3 山田太郎 03-1234-5678',from:'サンプル商事',subject:'書類'});
  const x=await xmlOf(d.buffer);assert.equal(d.name,'配送ラベル.docx');assert.match(x,/w:w="5669" w:h="8504"/);for(const s of ['お届け先','山田太郎　様','TEL：03-1234-5678','ご依頼主','サンプル商事','品名・備考','書類'])assert.ok(x.includes(s),s);
  const z=await xmlOf((await createPrintDocument({kind:'shipping',lang:'zh',paper:'a4',addresses:'100001 北京市示例路 张三 13800000000'})).buffer);assert.match(z,/w:val="dashed"/);assert.ok(z.includes('收件人')&&z.includes('电话：13800000000'));
});
test('模板库：例文接口只返回文字；寄件信息单可以生成文件，不覆盖同名',async t=>{
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'lib-'));t.after(()=>fs.rm(dir,{recursive:true,force:true}));
  const lib=createTemplateLibrary({root:process.cwd(),outputDir:async()=>dir,openFile:async()=>{}});
  assert.ok((await lib.run({action:'greeting'})).items.length>=4);assert.match((await lib.run({action:'greeting',id:'notice'})).text,/^拝啓/);
  const a=await lib.run({action:'document',kind:'shipping',lang:'zh',paper:'a4',addresses:'100001 北京市示例路 张三 13800000000'}),b=await lib.run({action:'document',kind:'shipping',lang:'zh',paper:'a4',addresses:'100001 北京市示例路 张三 13800000000'});
  assert.notEqual(a.path,b.path);assert.match(path.basename(b.path),/寄件信息单-2\.docx$/);
});
