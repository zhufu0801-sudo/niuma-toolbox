// 2026-11-25 日文文体转换（です・ます体 ⇄ である体），kuromoji 离线分词，不用 AI。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';
import {convertParagraphs,applyEdits,conjugate} from '../lib/ja-style.mjs';
async function conv(list,to,plain){const r=await convertParagraphs({paragraphs:list,to,plain});return list.map((t,i)=>applyEdits(t,r.paragraphs[i].edits));}

test('动词活用：五段各行、一段、サ変、カ変、ある',()=>{
  assert.equal(conjugate('書く','五段・カ行イ音便','ta'),'書いた');assert.equal(conjugate('行く','五段・カ行促音便','ta'),'行った');assert.equal(conjugate('泳ぐ','五段・ガ行','ta'),'泳いだ');
  assert.equal(conjugate('読む','五段・マ行','nai'),'読まない');assert.equal(conjugate('買う','五段・ワ行促音便','nai'),'買わない');assert.equal(conjugate('ある','五段・ラ行','nai'),'ない');
  assert.equal(conjugate('食べる','一段','renyo'),'食べ');assert.equal(conjugate('する','サ変・スル','vol'),'しよう');assert.equal(conjugate('来る','カ変・来ル','renyo'),'来');assert.equal(conjugate('くださる','五段・ラ行特殊','renyo'),'ください');
});
test('です・ます体 → である体：常见句末都能改，引用不动，请求句只提示',async()=>{
  const src=['資料をお送りします。','本日は会議です。','明日は雨でしょう。','確認しませんでした。','一緒に行きましょう。','本を借ります。家に帰ります。','担当者がおります。','こちらでございます。','問題はありません。','とても高いです。','本当ですか。','日程が決まりましたので、ご連絡いたします。','彼は「明日行きます」と言いました。','書きました。泳ぎました。死にました。'];
  assert.deepEqual(await conv(src,'plain'),['資料をお送りする。','本日は会議である。','明日は雨であろう。','確認しなかった。','一緒に行こう。','本を借りる。家に帰る。','担当者がいる。','こちらである。','問題はない。','とても高い。','本当か。','日程が決まったので、ご連絡する。','彼は「明日行きます」と言った。','書いた。泳いだ。死んだ。']);
  assert.deepEqual(await conv(['本日は会議です。','雨でした。'],'plain','da'),['本日は会議だ。','雨だった。']);
  const r=await convertParagraphs({paragraphs:['詳細は別紙をご参照ください。'],to:'plain'});assert.equal(r.edits,0);assert.equal(r.notes,1);
});
test('である体 → です・ます体：只改句末和“が、”“ので、”前的谓语，名词结尾和连体修饰不动',async()=>{
  const src=['資料を送る。','本日は会議である。','明日は雨だろう。','確認しなかった。','一緒に行こう。','問題はない。','昨日届いた。','とても高い。','行きたい。','安くなかった。','これは便利なのだ。','日程が決まったので、連絡する。','雨だが、出発する。','これは本ではない。','検討している。','結果は良好であろう。','本当か。','人が来る時間','彼は「明日行く」と言った。'];
  assert.deepEqual(await conv(src,'polite'),['資料を送ります。','本日は会議です。','明日は雨でしょう。','確認しませんでした。','一緒に行きましょう。','問題はありません。','昨日届きました。','とても高いです。','行きたいです。','安くありませんでした。','これは便利なのです。','日程が決まりましたので、連絡します。','雨ですが、出発します。','これは本ではありません。','検討しています。','結果は良好でしょう。','本当ですか。','人が来る時間','彼は「明日行く」と言いました。']);
  const r=await convertParagraphs({paragraphs:['中文段落，没有假名。',''],to:'polite'});assert.equal(r.edits,0);
});
test('文体转换：服务、功能区、Office 版和日文词条都接上了',()=>{
  const read=f=>fs.readFileSync(new URL('../'+f,import.meta.url),'utf8');
  assert.match(read('server.mjs'),/'ja-style'[,\]]/);assert.match(read('public/ribbon.xml'),/id="jaStyle"/);assert.match(read('public/tools/proof.js'),/id:'jaStyle'/);
  assert.match(read('public/office/office-batch2.js'),/jaStyle:jaStyle/);assert.match(read('lib/office-server.mjs'),/'\/api\/tools\/ja-style'/);assert.doesNotMatch(read('package.json'),/"kuromoji"/);for(const f of ['lib/vendor/kuromoji.min.cjs','lib/vendor/kuromoji-dict/base.dat.gz','lib/vendor/kuromoji-LICENSE.md'])assert.ok(fs.existsSync(new URL('../'+f,import.meta.url)),f);assert.match(read('scripts/licenses/build-notices.mjs'),/kuromoji-LICENSE\.md/);
  assert.match(read('scripts/package/build-release.mjs'),/pkg\.dependencies/);assert.match(read('public/i18n-ja-extra.js'),/"日文文体转换":/);
});
