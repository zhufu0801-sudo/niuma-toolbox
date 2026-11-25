// 2026-11-25 宏代码检查（不用 AI）和 住所 ↔ 郵便番号。
import test from 'node:test';import assert from 'node:assert/strict';
import fs from 'node:fs';import fsp from 'node:fs/promises';import os from 'node:os';import path from 'node:path';import vm from 'node:vm';import zlib from 'node:zlib';
import {buildFromCsv,createPostal,unzipFirst,cleanTown,createPostalStore,writeData,dataFromFile,downloadOfficial,findZipLink,OFFICIAL_URL,OFFICIAL_PAGE} from '../lib/postal-jp.mjs';
import {createCoachServer} from '../server.mjs';
const read=f=>fs.readFileSync(new URL('../'+f,import.meta.url),'utf8');
const same=(a,b)=>assert.equal(JSON.stringify(a),JSON.stringify(b));
function load(file,name){const ctx={};ctx.window=ctx;vm.runInNewContext(read(file),ctx);return ctx[name];}

// ---------- 宏代码检查 ----------
const V=load('public/vba-lint.js','VbaLint');
const codes=src=>V.lint(src).findings.map(f=>f.sev+':'+f.code+'@'+f.line);
test('宏检查：块没写完、Next 变量不对、Else 没有 If、多余的 End If 都能找到',()=>{
  const r=codes('Sub A()\n  If x Then\n    y = 1\n  For i = 1 To 3\n  Next j\nEnd Sub\nSub B()\n  End If\nEnd Sub\nFunction C()\n  With ActiveSheet\nEnd Function');
  assert.ok(r.includes('error:NEXTVAR@5'));assert.ok(r.includes('error:BLOCK@6'));assert.ok(r.includes('error:BLOCK@8'));assert.ok(r.includes('error:BLOCK@12'));
  same(codes('Option Explicit\nSub A()\n  Dim s As String: s = "If x Then"\n  If s = "" Then Exit Sub\n  For Each c In Range("A1:A3")\n    c.Value = 1\n  Next\n  Select Case s\n  Case "a"\n  End Select\nEnd Sub'),[]);
});
test('宏检查：单行 If、续行、注释和字符串里的关键字不算',()=>{
  same(codes("Option Explicit\nSub A()\n  If a Then b = 1 Else b = 2\n  ' If x Then\n  Rem For i = 1 To 2\n  MsgBox \"End If ' Next\"\n  s = \"a\" & _\n      \"b\"\n  Debug.Print s: If s = \"\" Then Exit Sub\nEnd Sub"),[]);
  assert.ok(codes('Option Explicit\nSub A()\n  s = "abc\nEnd Sub').includes('error:QUOTE@3'));
  assert.ok(codes('Option Explicit\nSub A()\n  x = (1 + 2\nEnd Sub').includes('error:PAREN@3'));
});
test('宏检查：跳转标签、Exit 写错、Set、Null 比较、PtrSafe',()=>{
  const r=codes('Option Explicit\nPrivate Declare Function GetTickCount Lib "kernel32" () As Long\nSub A()\n  On Error GoTo EH\n  GoTo Nowhere\n  Set o = CreateObject("x")\n  o2 = CreateObject("x")\n  If v = Null Then Exit Function\n  Exit Sub\nEH:\n  On Error GoTo 0\nEnd Sub');
  assert.ok(r.includes('error:PTRSAFE@2'));assert.ok(r.includes('error:LABEL@5'));assert.ok(!r.some(x=>x.startsWith('error:LABEL@4')));assert.ok(r.includes('error:SET@7'));assert.ok(!r.includes('error:SET@6'));assert.ok(r.includes('error:NULL@8'));assert.ok(r.includes('error:EXIT@8'));
});
test('宏检查：正着循环删行、设置没恢复、Resume Next、死循环、写死行号、Integer',()=>{
  const r=codes('Sub A()\n  Dim i As Integer\n  Application.ScreenUpdating = False\n  Application.Calculation = xlCalculationManual\n  On Error Resume Next\n  For i = 1 To 65536\n    If Cells(i, 1) = "" Then Rows(i).Delete\n  Next i\n  Do\n    DoEvents\n  Loop\nEnd Sub\nSub B()\n  Dim r As Long\n  For r = 100 To 1 Step -1\n    If Cells(r, 1) = "" Then Rows(r).EntireRow.Delete\n  Next r\nEnd Sub');
  for(const k of ['warn:DELETE_ASC@6','warn:RESTORE@3','warn:RESTORE@4','warn:ERR_RESUME@5','warn:INFINITE@9','tip:HARDROW@6','tip:INTEGER@2','tip:EXPLICIT@1'])assert.ok(r.includes(k),k);
  assert.ok(!r.some(x=>x.startsWith('warn:DELETE_ASC@15')));
});
test('宏检查：安全风险（自动运行、Shell、删文件、下载、改注册表、改宏代码、混淆）',()=>{
  const r=codes('Option Explicit\nSub Workbook_Open()\n  Shell "cmd /c x"\n  Kill "C:\\a.txt"\n  Set h = CreateObject("MSXML2.XMLHTTP")\n  CreateObject("WScript.Shell").RegWrite "HKCU\\x", 1\n  ThisWorkbook.VBProject.VBComponents(1).CodeModule.AddFromString "x"\n  s = Chr(65) & Chr(66) & Chr(67) & Chr(68) & Chr(69) & Chr(70) & Chr(71) & Chr(72)\nEnd Sub');
  for(const k of ['risk:AUTORUN@2','risk:SHELL@3','risk:DELETE@4','risk:NET@5','risk:REG@6','risk:SELFMOD@7','risk:OBFUSCATE@8'])assert.ok(r.includes(k),k);
  // 每条都有中日文说明
  for(const f of V.lint('Sub Auto_Open()\nShell "x"\nEnd Sub').findings){assert.ok(f.msg&&f.msgJa&&/[ぁ-ん]/.test(f.msgJa),f.code);}
});
test('宏检查：自动修改只改确定的（Integer→Long、Option Explicit、PtrSafe），字符串不动',()=>{
  const r=V.autofix('Attribute VB_Name = "Module1"\nDeclare Function F Lib "x" () As Integer\nSub A()\n  Dim i As Integer\n  MsgBox "As Integer"\nEnd Sub');
  const lines=r.code.split('\r\n');assert.equal(lines[0],'Attribute VB_Name = "Module1"');assert.equal(lines[1],'Option Explicit');assert.equal(lines[2],'Declare PtrSafe Function F Lib "x" () As Long');assert.equal(lines[4],'  Dim i As Long');assert.equal(lines[5],'  MsgBox "As Integer"');assert.equal(r.changes.length,3);
});

// ---------- 住所 ↔ 郵便番号 ----------
const L=(zip,pref,city,town)=>`99999,"${zip.slice(0,3)}  ","${zip}","カナ","カナ","カナ","${pref}","${city}","${town}",0,0,0,0,0,0`;
function sample(){const a=[L('1000000','東京都','千代田区','以下に掲載がない場合'),L('1000001','東京都','千代田区','千代田'),L('1000005','東京都','千代田区','丸の内（次のビルを除く）'),L('1006890','東京都','千代田区','丸の内ＪＰタワー（地階・階層不明）'),L('1006801','東京都','千代田区','丸の内ＪＰタワー（１階）'),L('1020072','東京都','千代田区','飯田橋'),L('1000003','東京都','千代田区','一ツ橋（１丁目）'),L('1010003','東京都','千代田区','一ツ橋（２丁目）'),
  L('0600042','北海道','札幌市中央区','大通西（１〜１９丁目）'),L('0640820','北海道','札幌市中央区','大通西（２０〜２８丁目）'),L('1900181','東京都','西多摩郡日の出町','大久野'),L('1900100','東京都','西多摩郡日の出町','以下に掲載がない場合'),L('1000301','東京都','利島村','利島村一円'),L('1001211','東京都','三宅島三宅村','坪田'),L('6008001','京都府','京都市下京区','四条通'),L('9638001','福島県','郡山市','大町'),L('3300063','埼玉県','さいたま市浦和区','高砂'),L('5410041','大阪府','大阪市中央区','北浜'),L('9200000','石川県','金沢市','以下に掲載がない場合'),L('9200001','石川県','金沢市','ケ丘町')];
  for(let i=0;i<90;i++)a.push(L(String(9990000+i),'沖縄県','テスト市','テスト町'+i));return '\ufeff'+a.join('\r\n')+'\r\n';}
function zipOf(text,dos=0){const raw=Buffer.from(text),comp=zlib.deflateRawSync(raw),name=Buffer.from('utf_ken_all.csv');
  const lh=Buffer.alloc(30);lh.writeUInt32LE(0x04034b50,0);lh.writeUInt16LE(8,8);lh.writeUInt32LE(comp.length,18);lh.writeUInt32LE(raw.length,22);lh.writeUInt16LE(name.length,26);
  const cd=Buffer.alloc(46);cd.writeUInt32LE(0x02014b50,0);cd.writeUInt16LE(8,10);cd.writeUInt16LE(dos,14);cd.writeUInt32LE(comp.length,20);cd.writeUInt32LE(raw.length,24);cd.writeUInt16LE(name.length,28);
  const body=Buffer.concat([lh,name,comp]),eo=Buffer.alloc(22);eo.writeUInt32LE(0x06054b50,0);eo.writeUInt16LE(1,8);eo.writeUInt16LE(1,10);eo.writeUInt32LE(46+name.length,12);eo.writeUInt32LE(body.length,16);return Buffer.concat([body,cd,name,eo]);}
test('邮编数据：官方 CSV 的特殊写法都处理了（以下に掲載がない場合、一円、（次のビルを除く）、ビルの階）',()=>{
  same(cleanTown('以下に掲載がない場合','千代田区'),{town:'',note:''});same(cleanTown('利島村一円','利島村'),{town:'',note:''});same(cleanTown('丸の内（次のビルを除く）','千代田区'),{town:'丸の内',note:''});
  assert.equal(cleanTown('丸の内ＪＰタワー（１階）','千代田区').building,true);same(cleanTown('大通西（１〜１９丁目）','札幌市中央区'),{town:'大通西',note:'１〜１９丁目',building:false});
  assert.throws(()=>buildFromCsv('a,b,c\n'),/不是日本邮政/);
});
test('邮编 → 住所、住所 → 邮编：郡名岛名可省略，丁目缩小候选，只认到市区町村时说明',()=>{
  const P=createPostal(buildFromCsv(sample(),{date:'2026-10-31'}));
  same(P.byZip('〒１００－０００１').candidates.map(c=>c.address),['東京都千代田区千代田']);assert.equal(P.byZip('12').reason,'format');assert.equal(P.byZip('9999999').reason,'notfound');
  const A=a=>P.byAddress(a).candidates.map(c=>c.zip+(c.level==='city'?'(市)':''));
  same(A('東京都千代田区千代田1-1'),['100-0001']);same(A('千代田区飯田橋３丁目'),['102-0072']);same(A('札幌市中央区大通西5丁目'),['060-0042']);same(A('札幌市中央区大通西二十二丁目'),['064-0820']);
  same(A('千代田区一ツ橋2-1'),['101-0003']);same(A('東京都日の出町大久野123'),['190-0181']);same(A('東京都西多摩郡日の出町平井'),['190-0100(市)']);same(A('三宅村坪田'),['100-1211']);same(A('東京都利島村248'),['100-0301(市)']);
  same(A('千代田区丸の内2-7-2'),['100-0005']);same(A('金沢市ヶ丘町1'),['920-0001']);assert.equal(P.byAddress('東京都存在しない区').ok,false);
  // 丁目で絞れないときは候選を全部返す
  assert.equal(P.byAddress('札幌市中央区大通西').candidates.length,2);
});
test('邮编数据：zip 解压、安装包自带和更新的数据用较新的那份、坏文件给出说明',async t=>{
  const d=await fsp.mkdtemp(path.join(os.tmpdir(),'postal-'));t.after(()=>fsp.rm(d,{recursive:true,force:true}));
  assert.equal(unzipFirst(zipOf(sample())).data.toString().length,sample().length);
  await fsp.writeFile(path.join(d,'utf_ken_all.zip'),zipOf(sample()));const data=await dataFromFile(path.join(d,'utf_ken_all.zip'),{date:'2026-11-01'});assert.ok(data.rows.length>100);
  const bundled=path.join(d,'bundled.json.gz'),user=path.join(d,'user.json.gz');await writeData({...data,date:'2026-09-30'},bundled);
  const store=createPostalStore({userFile:user,bundled});assert.equal((await store.status()).from,'bundled');
  await store.importFile(path.join(d,'utf_ken_all.zip'));const st=await store.status();assert.equal(st.from,'user');
  same((await store.zip(['1000001']))[0].candidates.map(c=>c.city),['千代田区']);
  await fsp.writeFile(user,'broken');const broken=createPostalStore({userFile:user,bundled});assert.match((await broken.status()).error,/损坏/);
  const none=createPostalStore({userFile:path.join(d,'x.gz'),bundled:path.join(d,'y.gz')});assert.match((await none.status()).error,/没有邮编数据/);
});
test('邮编接口：要本机口令；批量查；更新失败给出说明',async t=>{
  const d=await fsp.mkdtemp(path.join(os.tmpdir(),'postal-s-'));t.after(()=>fsp.rm(d,{recursive:true,force:true}));
  const bundled=path.join(d,'b.json.gz');await writeData(buildFromCsv(sample(),{date:'2026-10-31'}),bundled);
  const server=createCoachServer({cardsFile:path.join(d,'c.json'),pasteBarFile:path.join(d,'p.json'),settingsFile:path.join(d,'s.json'),backupDir:d,aiFile:path.join(d,'ai.json'),maskFile:path.join(d,'m.json'),postalFile:path.join(d,'u.json.gz'),postalOptions:{bundled}});
  await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>server.close(r)));
  const base=`http://127.0.0.1:${server.address().port}`,token=(await (await fetch(base+'/api/session')).json()).token;
  const post=(b,k=token)=>fetch(base+'/api/tools/postal',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':k},body:JSON.stringify(b)});
  assert.equal((await post({action:'status'},'bad')).status,403);
  const st=await (await post({action:'status'})).json();assert.equal(st.date,'2026-10-31');
  const r=await (await post({action:'address',items:['千代田区飯田橋1','','札幌市中央区大通西']})).json();same(r.results.map(x=>x.candidates.length),[1,0,2]);assert.equal(r.date,'2026-10-31');
  assert.equal((await post({action:'import',path:'relative.zip'})).status,400);
});
test('邮编批量补全：候选、只到市区町村、没找到分别标记',()=>{
  const P=load('public/postal-format.js','OfficePostal');
  assert.equal(P.looksZip('〒100-0001'),true);assert.equal(P.looksZip('１０００００１'),true);assert.equal(P.looksZip('東京都'),false);
  same(P.render({ok:true,candidates:[{zip:'100-0001',address:'東京都千代田区千代田',town:'千代田',level:'town'}]},'addr2zip'),{t:'100-0001',kind:'ok'});
  assert.equal(P.render({ok:true,candidates:[{zip:'190-0100',address:'x',town:'',level:'city'}]},'addr2zip').kind,'city');
  assert.equal(P.render({ok:true,candidates:[{zip:'060-0042',town:'大通西',note:'１〜１９丁目'},{zip:'064-0820',town:'大通西',note:'２０〜２８丁目'}]},'addr2zip').kind,'multi');
  same(P.render({ok:true,candidates:[{address:'A',town:'a'},{address:'B',town:'b'}]},'zip2addr'),{t:'候选：A / B',kind:'multi'});same(P.render({ok:false,reason:'format'},'zip2addr'),{t:'',kind:'format'});
});
test('宏代码检查和邮编：功能区、侧栏、Office 版、日文词条、安装包里的数据都接上了',()=>{
  const x=read('public/ribbon.xml');assert.match(x,/id="macroCheck"/);assert.match(x,/id="postalFill"/);assert.match(x,/id="postalOpen"/);
  for(const f of ['public/vba-lint.js','public/vba-panel.js','public/postal-panel.js','public/postal-format.js','public/tools/postal.js','public/icons/macrocheck.svg','public/icons/postal.svg','scripts/data/update-postal-jp.mjs'])assert.ok(fs.existsSync(new URL('../'+f,import.meta.url)),f);
  assert.match(read('public/tools.html'),/vba-panel\.js[\s\S]*postal-panel\.js/);assert.match(read('public/index.html'),/postal-format\.js"><\/script><script src="tools\/postal\.js/);assert.match(read('public/office/taskpane.html'),/\.\.\/postal-format\.js[\s\S]*office-batch2\.js[\s\S]*\.\.\/postal-panel\.js/);
  assert.match(read('lib/office-server.mjs'),/'\/vba-lint\.js','\/vba-panel\.js','\/postal-panel\.js','\/postal-format\.js'/);assert.match(read('public/main.js'),/postalOpen:'jumpPostal',macroCheck:'jumpMacroCheck'/);assert.match(read('public/main.js'),/FORM_FIRST=\{postalFill:1/);
  assert.match(read('public/office/office-batch2.js'),/TOOLS\.excel\.postalFill=postalFill/);for(const k of [/"宏代码检查":\s*"マクロのチェック"/,/"邮编 ↔ 住所":\s*"郵便番号 ↔ 住所"/])assert.match(read('public/i18n-ja-extra.js'),k);
});

// 2026-11-25：官网旧地址 404；数据日期曾经显示成导入那天。
const dos=(y,m,d)=>((y-1980)<<9)|(m<<5)|d;
test('邮编数据：数据日期取 zip 里的文件日期，不是导入那天',async()=>{
  assert.equal(unzipFirst(zipOf(sample(),dos(2026,10,30))).date,'2026-10-30');assert.equal(unzipFirst(zipOf(sample())).date,'');
  const d=await fsp.mkdtemp(path.join(os.tmpdir(),'postal-date-')),f=path.join(d,'utf_ken_all.zip');
  await fsp.writeFile(f,zipOf(sample(),dos(2026,10,30)));assert.equal((await dataFromFile(f)).date,'2026-10-30');
  const user=path.join(d,'u.json.gz'),store=createPostalStore({userFile:user,bundled:path.join(d,'none.json.gz')});
  assert.equal((await store.importFile(f)).date,'2026-10-30');
});
test('邮编数据：用新网址下载；网址又变了就去下载页面找链接；不跟去别的网站',async()=>{
  assert.ok(OFFICIAL_URL.startsWith('https://www.post.japanpost.jp/service/search/zipcode/download/'));
  assert.equal(findZipLink('<a href="utf/zip/utf_ken_all.zip">','https://www.post.japanpost.jp/a/b/utf-zip.html'),'https://www.post.japanpost.jp/a/b/utf/zip/utf_ken_all.zip');
  assert.equal(findZipLink('<a href="https://evil.example/utf_ken_all.zip">'),'');assert.equal(findZipLink('<a href="http://www.post.japanpost.jp/utf_ken_all.zip">'),'');
  const zip=zipOf(sample(),dos(2026,10,30)),moved='https://www.post.japanpost.jp/new/place/utf_ken_all.zip',calls=[];
  const res=(ok,body,h={})=>({ok,status:ok?200:404,headers:{get:k=>h[k.toLowerCase()]||null},arrayBuffer:async()=>body.buffer.slice(body.byteOffset,body.byteOffset+body.length),text:async()=>String(body)});
  const fetchImpl=async u=>{calls.push(u);if(u===moved)return res(true,zip);if(u===OFFICIAL_PAGE)return res(true,Buffer.from('<a href="/new/place/utf_ken_all.zip">全国一括</a>'));return res(false,Buffer.alloc(0));};
  const data=await downloadOfficial({fetchImpl});assert.equal(data.date,'2026-10-30');assert.ok(data.rows.length>100);same(calls,[OFFICIAL_URL,OFFICIAL_PAGE,moved]);
  const direct=await downloadOfficial({fetchImpl:async u=>u===OFFICIAL_URL?res(true,zipOf(sample()),{'last-modified':'Fri, 30 Oct 2026 01:00:00 GMT'}):res(false,Buffer.alloc(0))});assert.equal(direct.date,'2026-10-30');
  await assert.rejects(downloadOfficial({fetchImpl:async()=>res(false,Buffer.alloc(0))}),/HTTP 404.*utf_ken_all\.zip 后导入/);
});
