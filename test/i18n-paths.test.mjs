import test from 'node:test';import assert from 'node:assert/strict';import vm from 'node:vm';import fs from 'node:fs';
const code=fs.readFileSync(new URL('../public/i18n.js',import.meta.url),'utf8');
function load(){const w={OfficeI18nJa:{'已备份到：':'バックアップしました：','办公工具箱':'オフィスツールボックス','报告':'レポート'},localStorage:{getItem:()=> 'ja',setItem(){}},navigator:{language:'ja'},setInterval(){},setTimeout(){}};w.window=w;vm.runInNewContext(code,w);return w.OfficeI18n;}
test('日文状态提示保留 Windows、UNC 和 Mac 路径原文',()=>{const I=load();for(const p of ['C:\\办公工具箱\\报告 1.docx','\\\\server\\报告.docx','/Users/test/办公工具箱/报告.docx'])assert.equal(I.t('已备份到：'+p),'バックアップしました：'+p);});

test('没有目录项的提示使用日文',()=>{const w={localStorage:{getItem:()=> 'ja'},navigator:{language:'ja'},setTimeout(){},setInterval(){}};w.window=w;for(const f of ['i18n-ja.js','i18n-ja-extra.js','i18n.js'])vm.runInNewContext(fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8'),w);assert.equal(w.OfficeI18n.t('没有找到目录项。'),'目次に追加できる見出しが見つかりませんでした。');assert.equal(w.OfficeI18n.t('未完成：未找到一至三级标题，请先把章节设为标题。\n如文档已有部分更改，可使用 WPS 自带的撤销按钮。'),'完了しませんでした：見出し 1～3 が見つかりません。先に章を見出しにしてください。\n途中まで変更された場合は、WPS の「元に戻す」ボタンが使えます。');});
test('Office 动态加入日文字典后，已缓存的分段翻译也能使用新句子',()=>{const I=load();I.t('已备份到：x');I.extend({'备份的是最近一次保存的内容。':'最後に保存した内容をバックアップしました。'});assert.equal(I.t('已备份到：C:\\报告.docx\n备份的是最近一次保存的内容。'),'バックアップしました：C:\\报告.docx\n最後に保存した内容をバックアップしました。');});
test('多行消息里的路径不会变成 undefined（真实词典，逐行再翻译）',()=>{
  const w={localStorage:{getItem:()=>'ja',setItem(){}},navigator:{language:'ja'},location:{search:'?lang=ja'},document:{documentElement:{},addEventListener(){},querySelectorAll:()=>[]},MutationObserver:function(){this.observe=()=>{};},setInterval(){},setTimeout(){}};w.window=w;
  for(const f of ['i18n-ja.js','i18n-ja-extra.js','i18n.js'])vm.runInNewContext(fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8'),w);
  const I=w.OfficeI18n;try{I.set('ja');}catch{}
  const out=I.t('已转换为 Word：C:\\Users\\a\\办公工具箱\\传真.docx\n共 1 页、6 段。');
  assert.equal(out,'Word に変換しました：C:\\Users\\a\\办公工具箱\\传真.docx\n全 1 ページ、6 段落。');
  assert.doesNotMatch(I.t('已把 3 张图片存到：D:\\图片\n找到 2 个版本（新的在前），存放在：\\\\srv\\备份'),/undefined/);
});

test('WPS 工具失败提示逐段翻译，日文错误不会夹带中文前后缀',async()=>{
  const els={};const get=id=>els[id]||(els[id]={value:'',className:'',classList:{add(){}}});
  const w={localStorage:{getItem:()=>'ja',setItem(){}},navigator:{language:'ja'},location:{search:'?lang=ja'},document:{documentElement:{},getElementById:get,addEventListener(){},querySelectorAll:()=>[]},MutationObserver:function(){this.observe=()=>{};},setInterval(){},setTimeout(){},ToolIcons:{path:x=>x},OfficeTools:{host:()=> 'et',app:()=>({ActiveWorkbook:{Name:'qa.xlsx',FullName:'qa.xlsx'}}),runFromUi:async()=>{throw Error('右隣の列にすでに内容があります。');}}};w.window=w;
  for(const f of ['i18n-ja.js','i18n-ja-extra.js','i18n.js'])vm.runInNewContext(fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8'),w);
  const src=fs.readFileSync(new URL('../public/tools-ui.js',import.meta.url),'utf8');vm.runInNewContext(src.slice(0,src.indexOf('  function setView'))+'window.testRun=run;})();',w);
  await w.testRun('idCheck');assert.match(els.status.textContent,/完了しませんでした：右隣/);assert.match(els.status.textContent,/元に戻せない/);assert.doesNotMatch(els.status.textContent,/未完成|如已出现|撤销/);
  assert.equal(w.OfficeI18n.t('（Office 版给出操作步骤）'),'（Office 版では操作手順を表示）');
});

test('删除空白页的新提示（现在页数 / 状态栏可能没刷新）有日文',()=>{
  const w={localStorage:{getItem:()=> 'ja'},navigator:{language:'ja'},setTimeout(){},setInterval(){}};w.window=w;for(const f of ['i18n-ja.js','i18n-ja-extra.js','i18n.js'])vm.runInNewContext(fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8'),w);
  const t=w.OfficeI18n.t;
  assert.equal(t('已删除 2 处空段落或分页符。现在共 1 页（原来 3 页）。WPS 底部状态栏的页数有时不会马上更新，以这里为准。'),'空の段落または改ページを 2 個削除しました。現在 1 ページです（削除前は 3 ページ）。WPS 下部のステータスバーのページ数はすぐに更新されないことがあります。こちらの数字が正しい値です。');
  const r=t('已删除 2 处空段落或分页符，并缩小了文末空段落。原来共 3 页，空白的是第 2、3 页。WPS 底部状态栏的页数有时不会马上更新，可以滚动到文末确认。');
  assert.match(r,/^空の段落または改ページを 2 個削除しました。文末の空の段落も最小にしました。削除前は全 3 ページで、空白は .+ページ目でした。WPS 下部のステータスバー.+文末までスクロールして確認してください。$/);assert.doesNotMatch(r,/[页状态栏确认]/);
});
