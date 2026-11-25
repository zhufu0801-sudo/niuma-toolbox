// 数据脱敏（2026-11-25）：词表方案、对照表（加密保存）、卡片 / 粘贴栏里的脱敏和还原。
// 识别和替换规则在 public/mask-engine.js（侧栏和这里共用一份）。
//
// 对照表 = “替换成了什么 ↔ 原文”。它本身就是敏感数据，所以：
//  - 只存在这台电脑的 generated/masking.json；内容用 AES-256-GCM 加密；
//  - 加密用的钥匙交给系统保管：Windows 用 DPAPI（当前 Windows 用户才能解开），Mac 放进钥匙串；
//    文件被复制到别的电脑也打不开。开发、测试环境没有这些时不加密，界面上会说明。
//  - 词表（客户名、项目名）也是敏感的，同样加密。
import fs from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';
import { createCipheriv, createDecipheriv, randomBytes } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { createSecretBox } from './secret.mjs';

const ENGINE_FILE = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'public', 'mask-engine.js');
let engine = null;
export async function maskEngine() {
  if (engine) return engine;
  const ctx = {}; ctx.window = ctx;
  vm.runInNewContext(await fs.readFile(ENGINE_FILE, 'utf8'), ctx, { filename: 'mask-engine.js' });
  return (engine = ctx.MaskEngine);
}

export const MODES = ['star', 'block', 'code'];
const RECORD_LIMIT = 300, PROFILE_LIMIT = 30;
const DEFAULT_PREFS = { mode: 'star', ai: true, agent: true, autoModules: false, types: ['WORD', 'ID', 'CARD', 'MYNUMBER', 'PHONE', 'EMAIL'] };

function seal(key, value) {
  const iv = randomBytes(12), c = createCipheriv('aes-256-gcm', key, iv);
  const data = Buffer.concat([c.update(Buffer.from(JSON.stringify(value), 'utf8')), c.final()]);
  return { iv: iv.toString('base64'), tag: c.getAuthTag().toString('base64'), data: data.toString('base64') };
}
function open(key, box) {
  if (!box) return null;
  const d = createDecipheriv('aes-256-gcm', key, Buffer.from(box.iv, 'base64'));
  d.setAuthTag(Buffer.from(box.tag, 'base64'));
  return JSON.parse(Buffer.concat([d.update(Buffer.from(box.data, 'base64')), d.final()]).toString('utf8'));
}
const clean = (s, n) => String(s ?? '').replace(/[\x00-\x1f]/g, ' ').slice(0, n);
function cleanPrefs(p = {}, old = DEFAULT_PREFS) {
  const types = Array.isArray(p.types) ? p.types.filter(t => DEFAULT_PREFS.types.includes(t)) : old.types;
  return { agent: p.agent === undefined ? old.agent !== false : !!p.agent, mode: MODES.includes(p.mode) ? p.mode : old.mode, ai: p.ai === undefined ? old.ai : !!p.ai, autoModules: p.autoModules === undefined ? old.autoModules : !!p.autoModules, types: types.length ? types : old.types };
}
function cleanProfiles(list) {
  if (!Array.isArray(list)) throw Error('词表方案格式不对。');
  if (list.length > PROFILE_LIMIT) throw Error(`词表方案最多 ${PROFILE_LIMIT} 套。`);
  return list.map((p, i) => {
    const words = String(p && p.words || '');
    if (words.length > 50000) throw Error('一套词表最多 5 万字。');
    return { id: typeof (p && p.id) === 'string' && /^[a-z0-9]{4,24}$/.test(p.id) ? p.id : Date.now().toString(36) + i, name: clean(p && p.name, 40).trim() || `方案 ${i + 1}`, words };
  });
}

// file：保存位置；secrets：系统加密（测试时传假的）。
export function createMaskVault(file, { secrets = createSecretBox({ service: 'OfficeToolbox.masking-key', label: '办公工具箱 数据脱敏对照表' }) } = {}) {
  let state = null, key = null, queue = Promise.resolve();
  const serial = fn => { const r = queue.catch(() => {}).then(fn); queue = r; return r; };
  async function load() {
    if (state) return state;
    if (!file) { state = { v: 1, prefs: { ...DEFAULT_PREFS }, records: [] }; return state; }
    try { state = JSON.parse(await fs.readFile(file, 'utf8')); if (!state || state.v !== 1) throw Error('版本不对'); }
    catch (e) { if (e.code !== 'ENOENT') throw Error('对照表文件无法读取（' + e.message + '）。原文件已保留，请先备份 generated\\masking.json 再检查。'); state = { v: 1, prefs: { ...DEFAULT_PREFS }, records: [] }; }
    state.prefs = cleanPrefs(state.prefs || {}); state.records = Array.isArray(state.records) ? state.records : [];
    return state;
  }
  async function write() {
    if (!file) return;
    await fs.mkdir(path.dirname(file), { recursive: true });
    const temp = file + '.' + randomBytes(4).toString('hex') + '.tmp';
    await fs.writeFile(temp, JSON.stringify(state), 'utf8');
    await fs.rename(temp, file);
  }
  async function getKey() {
    if (key) return key;
    const s = await load();
    if (s.key && s.key.value) {
      let plain;
      try { plain = await secrets.unprotect(s.key.value, s.key.protection); }
      catch { throw Error('打不开对照表：加密钥匙只能在原来的电脑、原来的 Windows / Mac 用户下使用。'); }
      key = Buffer.from(plain, 'base64');
      if (key.length !== 32) throw Error('对照表的加密钥匙不完整，无法打开。');
      return key;
    }
    const raw = randomBytes(32);
    let sealed;
    try { sealed = await secrets.protect(raw.toString('base64')); } catch { sealed = { value: raw.toString('base64'), protection: 'none' }; }
    s.key = { value: sealed.value, protection: sealed.protection };
    key = raw; await write();
    return key;
  }
  const protection = async () => { await getKey(); return (await load()).key.protection; };
  const meta = r => ({ id: r.id, created: r.created, kind: r.kind, host: r.host, name: r.name, path: r.path, source: r.source, mode: r.mode, counts: r.counts, total: r.total, restored: r.restored || null });
  return {
    // 侧栏打开时读：偏好、词表方案、对照表列表（不含对照表内容）。
    state() { return serial(async () => { const s = await load(), k = await getKey(); return { prefs: s.prefs, profiles: open(k, s.profiles) || [], current: s.current || '', protection: s.key.protection, records: s.records.map(meta) }; }); },
    saveProfiles(input) { return serial(async () => {
      const s = await load(), k = await getKey(), list = cleanProfiles(input.profiles);
      s.profiles = seal(k, list); s.current = list.some(p => p.id === input.current) ? input.current : (list[0] ? list[0].id : '');
      if (input.prefs) s.prefs = cleanPrefs(input.prefs, s.prefs);
      await write(); return { profiles: list, current: s.current, prefs: s.prefs };
    }); },
    savePrefs(p) { return serial(async () => { const s = await load(); s.prefs = cleanPrefs(p, s.prefs); await write(); return s.prefs; }); },
    // 当前方案的词（AI 发送前隐藏、卡片自动脱敏用）。没有方案时是空的。
    currentWords() { return serial(async () => { const s = await load(); if (!s.profiles) return ''; const k = await getKey(), list = open(k, s.profiles) || [], p = list.find(x => x.id === s.current) || list[0]; return p ? p.words : ''; }); },
    prefs() { return serial(async () => (await load()).prefs); },
    // 存一份对照表。data 是侧栏交来的还原数据（各宿主格式不同，这里只加密保存）。
    addRecord(input) { return serial(async () => {
      const s = await load(), k = await getKey();
      const rec = { id: Date.now().toString(36) + randomBytes(3).toString('hex'), created: Date.now(), kind: ['doc', 'modules', 'clip'].includes(input.kind) ? input.kind : 'doc', host: ['wps', 'et', 'wpp', 'office-word', 'office-excel', 'modules'].includes(input.host) ? input.host : 'wps',
        name: clean(input.name, 120), path: clean(input.path, 500), source: clean(input.source, 500), mode: MODES.includes(input.mode) ? input.mode : 'star',
        counts: Object.fromEntries(Object.entries(input.counts || {}).filter(([t, n]) => /^[A-Z]{2,10}$/.test(t) && Number.isFinite(n)).slice(0, 10)), total: Number(input.total) || 0, data: seal(k, { pairs: input.pairs || [], data: input.data || null }) };
      s.records.unshift(rec);
      let clips=0;s.records=s.records.filter(r=>r.kind!=='clip'||++clips<=20);
      if (s.records.length > RECORD_LIMIT) s.records.length = RECORD_LIMIT;
      await write(); return meta(rec);
    }); },
    clipPairs() { return serial(async () => { const s=await load(), k=await getKey(); return s.records.filter(r=>r.kind==='clip').slice(0,20).flatMap(r=>open(k,r.data).pairs||[]); }); },
    getRecord(id) { return serial(async () => { const s = await load(), r = s.records.find(x => x.id === id); if (!r) throw Error('这份对照表已经不在了。'); return { ...meta(r), ...open(await getKey(), r.data) }; }); },
    // 按文件位置找：打开脱敏后的文件点“还原”时，自动找到它的对照表。
    findByPath(p) { return serial(async () => { const s = await load(), want = String(p || '').toLowerCase(); return want ? s.records.filter(r => r.path && r.path.toLowerCase() === want).map(meta) : []; }); },
    markRestored(id, info) { return serial(async () => { const s = await load(), r = s.records.find(x => x.id === id); if (r) { r.restored = { at: Date.now(), path: clean(info && info.path, 500) }; await write(); } return !!r; }); },
    removeRecord(id) { return serial(async () => { const s = await load(), n = s.records.length; s.records = s.records.filter(r => r.id !== id); if (n === s.records.length) return false; await write(); return true; }); },
    protection
  };
}

// 对照表导出成 CSV（Excel 能直接打开）：原文、替换为、类型、次数。文件里是原文，提醒客户用完删掉。
export async function exportRecordCsv(record, dir, { lang = 'zh' } = {}) {
  const E = await maskEngine(), ja = lang === 'ja';
  const rows = [[ja ? '元の内容' : '原文', ja ? '置き換え後' : '替换为', ja ? '種類' : '类型', ja ? '回数' : '次数']];
  for (const p of record.pairs || []) rows.push([p.original, p.masked, E.label(p.type, ja), String(p.count || 1)]);
  const cell = v => { let s = String(v ?? ''); if (/^[=+\-@\t\r]/.test(s)) s = "'" + s; return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const base = (ja ? 'マスキング対応表-' : '脱敏对照表-') + String(record.name || record.id).replace(/\.[A-Za-z0-9]{1,5}$/, '').replace(/[\\/:*?"<>|\x00-\x1f]/g, '_').slice(0, 60);
  await fs.mkdir(dir, { recursive: true });
  let target = path.join(dir, base + '.csv'), n = 1;
  for (;;) { try { await fs.writeFile(target, '﻿' + rows.map(r => r.map(cell).join(',')).join('\r\n'), { flag: 'wx' }); break; } catch (e) { if (e.code !== 'EEXIST' || n > 50) throw e; target = path.join(dir, `${base}-${++n}.csv`); } }
  return target;
}

// 卡片 / 粘贴栏一条内容（{type:'text',text} 或 {type:'table',rows}）脱敏。返回新内容、每处 [替换为, 原文, 类型]、是否有改动。
export async function maskBlock(block, { mode = 'star', lang = 'zh', types, words = '', prior = [] } = {}) {
  const E = await maskEngine(), parsed = E.parseWords(words).words, sess = E.session({ mode, lang, words: parsed, prior }), opts = { types, words: parsed };
  if (block.type === 'text') { const r = E.maskText(block.text, sess, opts); return { block: { type: 'text', text: r.text }, seq: r.seq, pairs: sess.pairs() }; }
  const seq = [], rows = block.rows.map(row => row.map(v => { const r = E.maskText(v, sess, opts); seq.push(...r.seq); return r.text; }));
  return { block: { type: 'table', rows }, seq, pairs: sess.pairs() };
}
export const blockText = b => b.type === 'text' ? b.text : b.rows.map(r => r.join('\t')).join('\n');

// 找出卡片和粘贴栏里有哪些要脱敏的内容（只读，不改）。
export async function scanModules({ cards, pasteBar, types, words = '' }) {
  const E = await maskEngine(), parsed = E.parseWords(words).words, out = [];
  const add = (kind, x) => { const hits = E.find(blockText(x.block), { types, words: parsed }); if (hits.length) out.push({ kind, id: x.id, title: x.title, count: hits.length, types: [...new Set(hits.map(h => h.type))], linked: !!x.source }); };
  for (const c of await cards.list()) add('card', c);
  for (const p of await pasteBar.list()) add('paste', p);
  return out;
}

// 把选中的卡片 / 粘贴栏条目脱敏（就地改，原文记进对照表，可以还原）。
export async function maskModules({ cards, pasteBar, vault, items, mode, lang, types, words, auto = false }) {
  const E = await maskEngine(), parsed = E.parseWords(words).words, sess = E.session({ mode, lang, words: parsed }), opts = { types, words: parsed };
  const changes = [], counts = {}, all = [];
  const pick = new Set((items || []).map(x => x.kind + ':' + x.id));
  const lists = { card: await cards.list(), paste: await pasteBar.list() };
  for (const kind of ['card', 'paste']) for (const x of lists[kind]) {
    if (!pick.has(kind + ':' + x.id)) continue;
    let block, seq = [];
    if (x.block.type === 'text') { const r = E.maskText(x.block.text, sess, opts); block = { type: 'text', text: r.text }; seq = r.seq; }
    else block = { type: 'table', rows: x.block.rows.map(row => row.map(v => { const r = E.maskText(v, sess, opts); seq.push(...r.seq); return r.text; })) };
    if (!seq.length) continue;
    await (kind === 'card' ? cards : pasteBar).replaceBlock(x.id, block);
    seq.forEach(s => { counts[s[2]] = (counts[s[2]] || 0) + 1; }); all.push(...seq);
    changes.push({ kind, id: x.id, title: x.title, before: x.block, after: block });
  }
  if (!changes.length) return { changed: 0, record: null };
  const total = Object.values(counts).reduce((a, b) => a + b, 0);
  const record = await vault.addRecord({ kind: 'modules', host: 'modules', name: auto ? '卡片 / 粘贴栏（自动脱敏）' : `卡片 / 粘贴栏 ${changes.length} 条`, mode, counts, total, pairs: E.table([all]), data: { changes } });
  return { changed: changes.length, total, record };
}

// 还原卡片 / 粘贴栏：内容还是脱敏后的样子才换回去；被改过或删掉的跳过并说明。
export async function restoreModules({ cards, pasteBar, vault, id }) {
  const rec = await vault.getRecord(id);
  if (rec.kind !== 'modules' || !rec.data) throw Error('这份对照表不是卡片 / 粘贴栏的。');
  const lists = { card: await cards.list(), paste: await pasteBar.list() }, same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  let done = 0; const skipped = [];
  for (const c of rec.data.changes || []) {
    const now = lists[c.kind].find(x => x.id === c.id);
    if (!now) { skipped.push(c.title + '（已删除）'); continue; }
    if (!same(now.block, c.after)) { skipped.push(c.title + '（后来改过）'); continue; }
    await (c.kind === 'card' ? cards : pasteBar).replaceBlock(c.id, c.before); done++;
  }
  if (done) await vault.markRestored(id, {});
  return { restored: done, skipped };
}
