// 2026-11-19 客户体验测试里发现的问题，修好后用这些测试防止回退。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import vm from 'node:vm';import path from 'node:path';
import {parseAddress,parseShipping} from '../lib/address-print.mjs';
import {createPrintDocument} from '../lib/print-tools.mjs';
import {greetingText} from '../lib/letter-templates.mjs';

test('宛名：姓と名の間に空白があっても 1 人の氏名として扱う',()=>{
  assert.deepEqual(parseAddress('〒150-0001 東京都渋谷区神宮前1-2-3 鈴木 一郎','ja'),{zip:'1500001',address:'東京都渋谷区神宮前1-2-3',name:'鈴木　一郎'});
  assert.equal(parseAddress('〒150-0001 東京都渋谷区神宮前1-2-3 中村 一郎 様','ja').name,'中村　一郎　様');
  // 部署名は住所側に残し、氏名と混ぜない
  assert.equal(parseAddress('〒150-0001 東京都渋谷区1-2-3 株式会社サンプル 営業部 鈴木一郎','ja').name,'鈴木一郎');
  // 番地のない区名＋氏名は区名を氏名にしない
  assert.equal(parseAddress('大阪府 大阪市 北区 山田太郎','ja').name,'山田太郎');
});
test('宛名：単独の敬称は前の語と一緒に、CSV のカンマも区切りとして扱う',()=>{
  assert.equal(parseAddress('〒150-0001 東京都渋谷区1-2-3 株式会社サンプル 御中','ja').name,'株式会社サンプル　御中');
  assert.deepEqual(parseAddress('150-0001,東京都渋谷区神宮前1-2-3,鈴木一郎','ja'),{zip:'1500001',address:'東京都渋谷区神宮前1-2-3',name:'鈴木一郎'});
  assert.equal(parseAddress('100001 北京市朝阳区示例路12号 王五','zh').name,'王五');
});
test('寄件信息：手机号带空格、姓名写在前面、单独的“收”',()=>{
  assert.deepEqual(parseShipping('北京市朝阳区示例路12号 张三 138 0000 0000','zh'),{zip:'',address:'北京市朝阳区示例路12号',name:'张三',tel:'13800000000'});
  assert.deepEqual(parseShipping('张三 13800000000 北京市朝阳区示例路12号','zh'),{zip:'',address:'北京市朝阳区示例路12号',name:'张三',tel:'13800000000'});
  assert.equal(parseShipping('上海市浦东新区示例路88号3楼 李四 收 021-12345678','zh').name,'李四 收');
});
test('竖排信件：文件名用表题，没有表题才叫ご案内',async()=>{
  const msg=greetingText('moving',new Date(2026,10,19));
  assert.equal((await createPrintDocument({kind:'vertical',lang:'ja',subject:'移転のご挨拶',message:msg})).name,'移転のご挨拶.docx');
  assert.equal((await createPrintDocument({kind:'vertical',lang:'ja',subject:'a/b:c',message:msg})).name,'abc.docx');
  assert.equal((await createPrintDocument({kind:'vertical',lang:'ja',message:msg})).name,'ご案内.docx');
});

const root=path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/,'$1')),'..');
function i18n(){const ctx={};ctx.window=ctx;ctx.document={documentElement:{},addEventListener(){},querySelectorAll:()=>[]};ctx.localStorage={getItem:()=>'ja',setItem(){}};ctx.navigator={language:'ja'};ctx.location={search:'?lang=ja'};ctx.setTimeout=()=>0;ctx.MutationObserver=function(){this.observe=()=>{};};
  for(const f of ['i18n-ja.js','i18n-ja-extra.js','i18n.js'])vm.runInNewContext(fs.readFileSync(path.join(root,'public',f),'utf8'),ctx);const I=ctx.OfficeI18n;try{I.set('ja');}catch{}return I;}
test('日文：示例路径整句翻译，模板语言与中文模板名不再半翻',()=>{
  const I=i18n();
  assert.equal(I.t('请输入完整路径，例如 C:\\Users\\你的用户名\\Documents\\备份。'),'フルパスを入力してください（例：C:\\Users\\ユーザー名\\Documents\\バックアップ）。');
  assert.equal(I.t('简体中文'),'中国語（簡体字）');assert.equal(I.t('合同常用格式'),'契約書の書式（中国語）');assert.equal(I.t('分类 / 筛选列'),'分類 / 絞り込む列');
  // 用户路径仍然原样保留
  assert.match(I.t('文件已保存到：C:\\Users\\sample-user\\Documents\\报价单.docx'),/C:\\Users\\sample-user\\Documents\\报价单\.docx$/);
});
test('i18n：多行输入框的 placeholder 也翻译（输入的文字不动）',()=>{
  const src=fs.readFileSync(path.join(root,'public/i18n.js'),'utf8');
  assert.match(src,/el\.tagName==='TEXTAREA'\?\(el\.parentNode/);
});
test('快捷工具搜索：比对译文和关键词，并能找到“文件”页签的打印功能',()=>{
  const src=fs.readFileSync(path.join(root,'public/quick-panel.js'),'utf8');
  assert.match(src,/score\(x,query\)>0/);assert.match(src,/function score\(x,q\)\{var h=hay\(x\)/);assert.match(src,/OfficeI18n\.t\(v\|\|''\)/);
  for(const k of ['挨拶状','はがき','送り状','印鑑','見積書'])assert.ok(src.includes(k),k);
  assert.match(src,/if\(selected\.kind==='jump'\)\{jump\(selected\);return;\}/);
});
test('模板面板：结果提示放在刚点的按钮下面，寄件人可以多行',()=>{
  const src=fs.readFileSync(path.join(root,'public/library-panel.js'),'utf8');
  assert.match(src,/button\.insertAdjacentElement\('afterend',msg\)/);
  assert.match(src,/<textarea id="printFrom" rows="2">/);
});
