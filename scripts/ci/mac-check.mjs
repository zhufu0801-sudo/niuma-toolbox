// GitHub 的 macOS 机器上检查装好的办公工具箱（先运行 scripts/ci/mac-install.sh）。
//   node scripts/ci/mac-check.mjs          检查安装结果
//   node scripts/ci/mac-check.mjs uninstall 检查卸载
// 这台机器上没有 WPS 和 Office，所以只能测“不用界面”的部分：本机服务、登记文件、钥匙串、路径、Office 侧栏网页、卸载。
// 功能区按钮、侧栏里的操作还是要在真的 Mac 上点一遍。
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import https from 'node:https';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';

const HOME = os.homedir(), APP = path.join(HOME, 'Library', 'Application Support', 'OfficeToolbox', 'app'), BASE = 'http://127.0.0.1:39127';
const PLIST = path.join(HOME, 'Library', 'LaunchAgents', 'com.officetoolbox.server.plist');
const run = (cmd, args) => new Promise((ok, bad) => execFile(cmd, args, { encoding: 'utf8', timeout: 120000 }, (e, out, err) => e ? bad(Object.assign(e, { stderr: err })) : ok(out)));
const sleep = ms => new Promise(r => setTimeout(r, ms));
const exists = p => fs.stat(p).then(() => true, () => false);
let failed = 0;
async function step(name, fn) { try { await fn(); console.log('✓ ' + name); } catch (e) { failed++; console.log('✗ ' + name + '\n   ' + String(e && e.stack || e).split('\n').slice(0, 4).join('\n   ')); } }
async function health(ms = 30000) { const end = Date.now() + ms; let last; while (Date.now() < end) { try { const r = await fetch(BASE + '/api/health'); if (r.ok) return r.json(); last = 'HTTP ' + r.status; } catch (e) { last = e.message; } await sleep(1000); } throw Error('本机服务没有响应：' + last); }
async function restart() { const uid = String(process.getuid()); await run('launchctl', ['kickstart', '-k', `gui/${uid}/com.officetoolbox.server`]); await sleep(2000); return health(); }
async function token() { return (await (await fetch(BASE + '/api/session')).json()).token; }
async function mask(body) { const r = await fetch(BASE + '/api/tools/mask', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Coach-Token': await token() }, body: JSON.stringify(body) }); const d = await r.json(); if (!r.ok) throw Error(d.error || 'HTTP ' + r.status); return d; }
function getHttps(url) { return new Promise((ok, bad) => https.get(url, { rejectUnauthorized: false }, res => { let b = ''; res.setEncoding('utf8'); res.on('data', c => { b += c; }); res.on('end', () => ok({ status: res.statusCode, body: b })); }).on('error', bad)); }
async function showLog() { try { const log = await fs.readFile(path.join(APP, 'generated', 'server.log'), 'utf8'); console.log('---- server.log（最后 40 行）----\n' + log.split('\n').slice(-40).join('\n')); } catch { console.log('（没有 server.log）'); } }

if (process.argv[2] === 'uninstall') {
  await step('卸载：运行 uninstall.command', async () => { const pkg = (await fs.readdir('/tmp/unzip-check')).find(n => /macOS/.test(n)); await run('bash', [path.join('/tmp/unzip-check', pkg, 'uninstall.command')]); });
  await step('卸载后：自动启动项删掉了、服务停了', async () => { assert.equal(await exists(PLIST), false); await sleep(2000); let up = true; try { await fetch(BASE + '/api/health'); } catch { up = false; } assert.equal(up, false, '服务还在运行'); });
  await step('卸载后：WPS 登记里没有办公工具箱了，设置保留', async () => { const xml = await fs.readFile(path.join(HOME, '.kingsoft', 'wps', 'jsaddons', 'publish.xml'), 'utf8'); assert.doesNotMatch(xml, /WpsCoach/); assert.ok(await exists(path.join(APP, 'generated'))); });
} else {
  await step('本机服务由 launchd 自动启动，版本号正确', async () => { const h = await health(); const pkg = JSON.parse(await fs.readFile(path.join(APP, 'package.json'), 'utf8')); assert.equal(h.version, pkg.version); assert.match(await fs.readFile(PLIST, 'utf8'), /com\.officetoolbox\.server/); });
  await step('WPS 登记：publish.xml 里有文字、表格、演示三条', async () => { const xml = await fs.readFile(path.join(HOME, '.kingsoft', 'wps', 'jsaddons', 'publish.xml'), 'utf8'); for (const n of ['WpsCoachWriter', 'WpsCoachSheets', 'WpsCoachSlides']) assert.match(xml, new RegExp('name="' + n + '"')); });
  await step('WPS 加载项文件都能取到（功能区、脚本、图标、侧栏）', async () => { for (const host of ['wps', 'et', 'wpp']) for (const f of ['index.html', 'main.js', 'ribbon.xml']) { const r = await fetch(`${BASE}/addon/${host}/${f}`); assert.equal(r.status, 200, host + '/' + f); } for (const f of ['/tools.html', '/mask-engine.js', '/mask-host.js', '/mask-panel.js', '/icons/shield.svg']) assert.equal((await fetch(BASE + f)).status, 200, f); });
  await step('日文分词（kuromoji 词典在安装包里）', async () => { const r = await fetch(BASE + '/api/tools/ja-style', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Coach-Token': await token() }, body: JSON.stringify({ paragraphs: ['本を読みます。'], to: 'plain' }) }); const d = await r.json(); assert.equal(r.status, 200, d.error); assert.ok(d.edits >= 1, JSON.stringify(d)); });
  await step('日本邮编数据在安装包里，能查', async () => { const t = await token(), r = await fetch(BASE + '/api/tools/postal', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-Coach-Token': t }, body: JSON.stringify({ action: 'zip', items: ['100-0001'] }) }); const d = await r.json(); assert.equal(r.status, 200, d.error); assert.equal(d.results[0].candidates[0].city, '千代田区'); });
  await step('钥匙串：对照表加密，重启服务后还能读出来', async () => {
    const st = await mask({ action: 'state' }); assert.equal(st.protection, 'keychain');
    await mask({ action: 'profiles', profiles: [{ name: 'CI', words: '张三' }] });
    const rec = (await mask({ action: 'save', record: { kind: 'doc', host: 'wps', name: 'CI.docx', mode: 'code', counts: { WORD: 1 }, total: 1, pairs: [{ original: '张三', masked: '【名称A】', type: 'WORD', count: 1 }], data: { seq: {} } } })).record;
    assert.doesNotMatch(await fs.readFile(path.join(APP, 'generated', 'masking.json'), 'utf8'), /张三/);
    await restart();
    assert.equal((await mask({ action: 'get', id: rec.id })).pairs[0].original, '张三');
    await mask({ action: 'delete', id: rec.id });
  });
  await step('中文和带空格的路径：副本文件名放在原文件旁边', async () => {
    const dir = path.join(HOME, 'Documents', '测试 文件夹 2026'); await fs.mkdir(dir, { recursive: true }); await fs.writeFile(path.join(dir, '合同 草稿.docx'), 'x');
    const t = await mask({ action: 'target', source: path.join(dir, '合同 草稿.docx'), name: '合同 草稿.docx' }); assert.equal(t.path, path.join(dir, '合同 草稿-脱敏.docx'));
    const j = await mask({ action: 'target', source: path.join(dir, '合同 草稿.docx'), name: '合同 草稿.docx', lang: 'ja' }); assert.equal(path.basename(j.path), '合同 草稿-マスキング.docx');
  });
  await step('Office 侧栏：生成本机证书、放入加载项清单，https 网页能打开', async () => {
    // 真实安装时 macOS 会要求输密码来信任证书；这里跳过“信任”这一步（GitHub 机器上没人输密码），其余照常。
    const fake = await fs.mkdtemp(path.join(os.tmpdir(), 'office-home-'));
    for (const a of ['Word', 'Excel', 'Powerpoint']) await fs.mkdir(path.join(fake, 'Library', 'Containers', 'com.microsoft.' + a, 'Data', 'Documents'), { recursive: true });
    const { registerOffice } = await import(path.join(APP, 'lib', 'mac-install.mjs'));
    const written = await registerOffice(APP, fake, (cmd, args) => cmd === 'security' ? Promise.resolve('') : run(cmd, args));
    assert.equal(written.length, 3); assert.match(await fs.readFile(written[0], 'utf8'), /<OfficeApp[\s\S]*https:\/\/localhost:39128/);
    await restart(); await sleep(1500);
    const page = await getHttps('https://localhost:39128/office/taskpane.html'); assert.equal(page.status, 200); assert.match(page.body, /mask-panel\.js/);
    for (const f of ['/mask-engine.js', '/mask-host.js', '/mask-panel.js']) assert.equal((await getHttps('https://localhost:39128' + f)).status, 200, f);
  });
  await step('设置页和功能列表（常用组的重复按钮不在列表里）', async () => { assert.equal((await fetch(BASE + '/settings.html')).status, 200); const r = await fetch(BASE + '/api/features', { headers: { 'X-Coach-Token': await token() } }); assert.equal(r.status, 200); const ids = (await r.json()).features.map(f => f.id); assert.ok(ids.includes('dataMask')); assert.ok(!ids.some(id => /^c[wep]_/.test(id))); });
}
if (failed) { await showLog(); console.log(`\n${failed} 项没通过。`); process.exit(1); }
console.log('\n全部通过。');
