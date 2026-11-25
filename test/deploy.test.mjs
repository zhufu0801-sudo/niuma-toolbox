import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createSettingsStore } from '../lib/settings.mjs';
import { createAiStore } from '../lib/ai.mjs';
import { slim } from '../scripts/package/slim.mjs';
const here = p => new URL('../' + p, import.meta.url);
test('企业预设示例：settings.json 和 ai.json 能被工具箱正常读取', async t => {
  const dir = await fsp.mkdtemp(path.join(os.tmpdir(), 'preset-')); t.after(() => fsp.rm(dir, { recursive: true, force: true }));
  for (const f of ['settings.json', 'ai.json']) await fsp.copyFile(here('scripts/package/preset-example/' + f), path.join(dir, f));
  const s = (await createSettingsStore(path.join(dir, 'settings.json'), { backupDir: path.join(dir, 'b') }).get()).settings;
  assert.equal(s.ui.language, 'ja'); assert.equal(s.text.font, '游明朝'); assert.equal(s.cards.limit, 40);
  const v = await createAiStore(path.join(dir, 'ai.json'), { secrets: { available: false, protect: async x => ({ value: x, protection: 'none' }), unprotect: async x => x } }).publicView();
  assert.equal(v.config.provider, 'custom'); assert.equal(v.config.privacy.mode, 'mask'); assert.equal(v.config.hasKey, false);
});
test('安装脚本：企业预设只在文件不存在时复制，支持无人值守参数', () => {
  const ps = fs.readFileSync(here('scripts/package/Initialize.ps1'), 'utf8');
  assert.match(ps, /\[switch\]\$NoLaunch, \[switch\]\$NoShortcut/);
  assert.match(ps, /if \(Test-Path -LiteralPath \$dst\) \{ Write-Host[^\n]*continue \}/);
  assert.match(ps, /ConvertFrom-Json \| Out-Null \} catch \{ throw/);
  assert.ok(ps.indexOf('Copy-Preset\r\n  if ($chosen') > 0);
});
test('启动脚本：每条报错都有中文和日文', () => {
  const ps = fs.readFileSync(here('Launch.ps1'), 'utf8');
  const keys = [...ps.matchAll(/\(T '([a-z]+)'\)/g)].map(m => m[1]);
  assert.ok(keys.length >= 8);
  for (const k of new Set(keys)) assert.match(ps, new RegExp(`\\n  ${k} +=  ?@\\('[^']+', '[^']*[ぁ-んァ-ン][^']*'\\)`), k);
  assert.doesNotMatch(ps.replace(/\$texts = @\{[\s\S]*?\n\}/, ''), /throw '[^']*[一-鿿]/);
});
test('安装包瘦身：删除源码映射、类型声明、说明文档和浏览器版构建，许可证保留', async t => {
  const dir = await fsp.mkdtemp(path.join(os.tmpdir(), 'slim-')); t.after(() => fsp.rm(dir, { recursive: true, force: true }));
  const put = async (p, c = 'x') => { await fsp.mkdir(path.dirname(path.join(dir, p)), { recursive: true }); await fsp.writeFile(path.join(dir, p), c); };
  for (const p of ['node_modules/a/index.js', 'node_modules/a/index.js.map', 'node_modules/a/index.d.ts', 'node_modules/a/README.md', 'node_modules/a/LICENSE', 'node_modules/a/package.json', 'node_modules/a/test/x.js',
    'node_modules/a/lib/test/keep.js', 'node_modules/pdfjs-dist/build/pdf.mjs', 'node_modules/pdfjs-dist/legacy/build/pdf.mjs', 'node_modules/@types/node/index.d.ts', 'node_modules/docx/dist/index.mjs', 'node_modules/docx/dist/index.iife.js']) await put(p);
  const r = slim(dir); assert.ok(r.files === 7 && r.bytes > 0);
  const exists = p => fs.existsSync(path.join(dir, p));
  for (const p of ['node_modules/a/index.js', 'node_modules/a/LICENSE', 'node_modules/a/lib/test/keep.js', 'node_modules/pdfjs-dist/legacy/build/pdf.mjs', 'node_modules/docx/dist/index.mjs']) assert.ok(exists(p), p);
  for (const p of ['node_modules/a/index.js.map', 'node_modules/a/index.d.ts', 'node_modules/a/README.md', 'node_modules/a/test', 'node_modules/pdfjs-dist/build', 'node_modules/@types', 'node_modules/docx/dist/index.iife.js']) assert.ok(!exists(p), p);
});
