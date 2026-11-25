// 隐私保护关卡：所有发给 AI 的内容都先经过这里，收到的回答也经过这里。
//  1. 内置保护：发送前把身份证号、手机号、邮箱、银行卡号、日本个人编号等换成占位符（如 [PHONE_1]），
//     收到回答后再换回原文。占位符和原文的对应关系只在这一次请求的内存里，不写文件、不写日志。
//  2. 插件接口：app\plugins 文件夹里的 *.mjs 会在第一次使用 AI 时加载（以 _ 开头的文件不加载）。
//     插件可以在发送前改写或拦截内容，在收到后还原或检查回答。说明见 plugins\README.md。
//  3. 出错即停：插件报错、超时或返回格式不对，一律不发送（宁可失败，不泄露）。
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

export const PRIVACY_MODES = ['mask', 'remind', 'off'];
export const PLUGIN_API_VERSION = 1;

function luhn(digits) {
  let sum = 0, dbl = false;
  for (let i = digits.length - 1; i >= 0; i--) { let d = digits.charCodeAt(i) - 48; if (dbl) { d *= 2; if (d > 9) d -= 9; } sum += d; dbl = !dbl; }
  return sum % 10 === 0;
}
// 日本个人编号（マイナンバー）12 位，最后一位是校验位
function myNumberOk(digits) {
  if (!/^\d{12}$/.test(digits)) return false;
  let sum = 0;
  for (let n = 1; n <= 11; n++) { const p = Number(digits[11 - n]), q = n <= 6 ? n + 1 : n - 5; sum += p * q; }
  const r = sum % 11, check = r <= 1 ? 0 : 11 - r;
  return check === Number(digits[11]);
}
function idCnOk(id) {
  const w = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2], c = '10X98765432';
  let s = 0; for (let i = 0; i < 17; i++) s += Number(id[i]) * w[i];
  return c[s % 11] === id[17].toUpperCase();
}
// 顺序有讲究：先认长号码（身份证、银行卡），再认手机号等短号码。
export const DETECTORS = [
  { type: 'ID', label: '身份证号', labelJa: '身分証番号', re: /(?<![0-9A-Za-z])[1-9]\d{5}(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx](?![0-9A-Za-z])/g, ok: m => idCnOk(m) },
  { type: 'CARD', label: '银行卡号', labelJa: 'カード番号', re: /(?<![\d-])\d{4}(?:[ -]?\d{4}){2,3}(?:[ -]?\d{1,3})?(?![\d-])/g, ok: m => { const d = m.replace(/\D/g, ''); return d.length >= 15 && d.length <= 19 && luhn(d); } },
  { type: 'MYNUMBER', label: '个人编号', labelJa: 'マイナンバー', re: /(?<![\d-])\d{4}[ -]?\d{4}[ -]?\d{4}(?![\d-])/g, ok: m => myNumberOk(m.replace(/\D/g, '')) },
  { type: 'PHONE', label: '手机号', labelJa: '電話番号', re: /(?<![\d+])(?:\+?86[ -]?)?1[3-9]\d[ -]?\d{4}[ -]?\d{4}(?!\d)|(?<![\d+])(?:\+81[ -]?|0)[789]0[ -]?\d{4}[ -]?\d{4}(?!\d)/g },
  { type: 'EMAIL', label: '邮箱', labelJa: 'メールアドレス', re: /[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}/g }
];

// 扫描并替换。返回替换后的文字、对应表和每类数量。同一个原文在一次请求里用同一个占位符。
// words：数据脱敏词表里的词（客户名、项目名等，2026-11-25），先于号码替换成 [WORD_1]；长词在前。
const escapeRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
export function maskText(text, state, words = []) {
  let out = String(text ?? '');
  if (words.length) out = out.replace(new RegExp(words.map(escapeRe).join('|'), 'g'), m => {
    if (!m) return m;
    const known = state.byValue.get(m);
    if (known) return known;
    state.count.WORD = (state.count.WORD || 0) + 1;
    const token = `[WORD_${state.count.WORD}]`;
    state.byValue.set(m, token); state.byToken.set(token, m);
    return token;
  });
  for (const d of DETECTORS) {
    out = out.replace(d.re, m => {
      if (d.ok && !d.ok(m)) return m;
      const known = state.byValue.get(m);
      if (known) return known;
      state.count[d.type] = (state.count[d.type] || 0) + 1;
      const token = `[${d.type}_${state.count[d.type]}]`;
      state.byValue.set(m, token); state.byToken.set(token, m);
      return token;
    });
  }
  return out;
}
const WORD_INFO = { type: 'WORD', label: '词表里的词', labelJa: '単語リストの語' };
export function findPersonalInfo(text, words = []) {
  const state = { byValue: new Map(), byToken: new Map(), count: {} };
  maskText(text, state, words);
  return [WORD_INFO, ...DETECTORS].filter(d => state.count[d.type]).map(d => ({ type: d.type, label: d.label, labelJa: d.labelJa, count: state.count[d.type] }));
}
// 词表文字 → 词的列表（一行一个，“原文=代号”只取原文，# 开头是备注），长词在前。
export function wordList(text) {
  const seen = new Set();
  return String(text || '').split(/\r?\n/).map(l => l.trim()).filter(l => l && !l.startsWith('#')).map(l => l.split(/[=＝]/)[0].trim()).filter(w => w && w.length <= 100 && !seen.has(w) && seen.add(w)).sort((a, b) => b.length - a.length).slice(0, 500);
}
export function restoreText(text, byToken) {
  let out = String(text ?? '');
  for (const [token, value] of byToken) out = out.split(token).join(value);
  return out;
}
const cloneMessages = list => (Array.isArray(list) ? list : []).map(m => ({ role: String(m.role || 'user'), content: String(m.content ?? '') }));
function validMessages(list) {
  return Array.isArray(list) && list.length > 0 && list.length <= 50 && list.every(m => m && ['system', 'user', 'assistant'].includes(m.role) && typeof m.content === 'string' && m.content.length <= 200000);
}
function blocked(message, extra = {}) { return Object.assign(Error(message), { status: 451, privacy: true, ...extra }); }
function withTimeout(promise, ms, name) {
  let t; return Promise.race([promise, new Promise((_, reject) => { t = setTimeout(() => reject(blocked(`隐私保护插件“${name}”处理超时，已停止发送。`)), ms); })]).finally(() => clearTimeout(t));
}

// 读取 plugins 文件夹。只认本文件夹里直接放的 .mjs 文件；格式不对的插件会被列为“未加载”，并让 AI 功能停用，避免以为有保护其实没有。
export async function loadPlugins(dir, { log = console } = {}) {
  const loaded = [], failed = [];
  let names = [];
  try { names = (await fs.readdir(dir)).filter(n => n.endsWith('.mjs') && !n.startsWith('_')).sort(); } catch { return { loaded, failed }; }
  for (const file of names) {
    try {
      const mod = await import(pathToFileURL(path.join(dir, file)).href);
      const p = mod.default || mod.plugin;
      if (!p || typeof p !== 'object' || typeof p.name !== 'string' || !p.name.trim()) throw Error('缺少 name');
      if (p.beforeSend && typeof p.beforeSend !== 'function') throw Error('beforeSend 必须是函数');
      if (p.afterReceive && typeof p.afterReceive !== 'function') throw Error('afterReceive 必须是函数');
      if (!p.beforeSend && !p.afterReceive) throw Error('至少要有 beforeSend 或 afterReceive');
      if (p.apiVersion !== undefined && p.apiVersion !== PLUGIN_API_VERSION) throw Error(`接口版本 ${p.apiVersion} 与工具箱（${PLUGIN_API_VERSION}）不一致`);
      loaded.push({ file, name: p.name.trim().slice(0, 60), version: String(p.version || '').slice(0, 20), plugin: p });
    } catch (e) {
      failed.push({ file, error: String(e && e.message || e).slice(0, 200) });
      try { log.error?.(`[privacy] 插件 ${file} 未加载：${e && e.message}`); } catch {}
    }
  }
  return { loaded, failed };
}

// mode：mask = 自动隐藏后发送；remind = 发现个人信息时先让用户确认；off = 不处理（插件仍会运行）。
// getWords：返回数据脱敏词表的文字（设置了“发给 AI 前也隐藏这些词”时才有）。
export function createPrivacyGuard({ getMode = async () => 'mask', getWords = async () => '', pluginDir = '', plugins = null, timeout = 5000, log = console } = {}) {
  let ready = null;
  function load() {
    if (!ready) ready = plugins ? Promise.resolve({ loaded: plugins.map((p, i) => ({ file: `inline-${i}`, name: p.name, version: p.version || '', plugin: p })), failed: [] }) : loadPlugins(pluginDir, { log });
    return ready;
  }
  async function info() {
    const { loaded, failed } = await load(), mode = await getMode();
    return { mode: PRIVACY_MODES.includes(mode) ? mode : 'mask', apiVersion: PLUGIN_API_VERSION, plugins: loaded.map(p => ({ name: p.name, version: p.version, file: p.file })), failed };
  }
  // ctx：{ messages, feature, provider, region, model, confirmed }
  async function beforeSend(ctx) {
    const { loaded, failed } = await load();
    if (failed.length) throw blocked(`隐私保护插件“${failed[0].file}”没有正确加载（${failed[0].error}），为安全起见已停止发送。请修正或移走该插件。`);
    let mode = await getMode(); if (!PRIVACY_MODES.includes(mode)) mode = 'mask';
    let messages = cloneMessages(ctx.messages);
    const state = { byValue: new Map(), byToken: new Map(), count: {} };
    let words = [];
    try { words = wordList(await getWords()); } catch (e) { throw blocked('读不出数据脱敏的词表（' + String(e && e.message || e).slice(0, 120) + '），为安全起见已停止发送。'); }
    const findings = findPersonalInfo(messages.map(m => m.content).join('\n'), words);
    if (mode === 'remind' && findings.length && !ctx.confirmed)
      throw Object.assign(blocked('要发送的内容里有个人信息：' + findings.map(f => `${f.label} ${f.count} 处`).join('、') + '。确定要原样发给 AI 吗？（在设置里选“自动隐藏”可以自动替换后再发送）'), { status: 428, needConfirm: true, findings });
    if (mode === 'mask' && findings.length) {
      messages = messages.map(m => ({ ...m, content: maskText(m.content, state, words) }));
      const note = '文中形如 [PHONE_1]、[EMAIL_1] 的方括号内容是隐私占位符，请在回答中原样保留，不要改写或猜测。 / [PHONE_1] などの角括弧の語は個人情報のプレースホルダーです。そのまま残してください。';
      const sys = messages.find(m => m.role === 'system');
      if (sys) sys.content += '\n' + note; else messages.unshift({ role: 'system', content: note });
    }
    const shared = { feature: String(ctx.feature || ''), provider: String(ctx.provider || ''), region: String(ctx.region || ''), model: String(ctx.model || ''), mode };
    const pluginState = new Map(), used = [];
    for (const p of loaded) {
      if (!p.plugin.beforeSend) continue;
      const scratch = {}; pluginState.set(p, scratch);
      let r;
      try { r = await withTimeout(Promise.resolve(p.plugin.beforeSend({ ...shared, messages: cloneMessages(messages), state: scratch })), timeout, p.name); }
      catch (e) { if (e && e.privacy) throw e; throw blocked(`隐私保护插件“${p.name}”出错，已停止发送：${String(e && e.message || e).slice(0, 160)}`); }
      if (r == null) { used.push(p.name); continue; }
      if (typeof r !== 'object') throw blocked(`隐私保护插件“${p.name}”返回格式不对，已停止发送。`);
      if (r.block) throw blocked(String(r.reason || `隐私保护插件“${p.name}”拦截了这次发送。`).slice(0, 300), { plugin: p.name });
      if (r.messages !== undefined) { if (!validMessages(r.messages)) throw blocked(`隐私保护插件“${p.name}”返回的内容格式不对，已停止发送。`); messages = cloneMessages(r.messages); }
      used.push(p.name);
    }
    const masked = Object.values(state.count).reduce((a, b) => a + b, 0);
    return { messages, session: { state, pluginState, shared, loaded }, report: { mode, masked, findings, plugins: used } };
  }
  async function afterReceive(session, answer) {
    let out = String(answer ?? '');
    for (const p of [...session.loaded].reverse()) {
      if (!p.plugin.afterReceive) continue;
      let r;
      try { r = await withTimeout(Promise.resolve(p.plugin.afterReceive({ ...session.shared, answer: out, state: session.pluginState.get(p) || {} })), timeout, p.name); }
      catch (e) { if (e && e.privacy) throw e; throw blocked(`隐私保护插件“${p.name}”处理回答时出错：${String(e && e.message || e).slice(0, 160)}`); }
      if (r == null) continue;
      if (typeof r !== 'object' || (r.answer !== undefined && typeof r.answer !== 'string')) throw blocked(`隐私保护插件“${p.name}”返回格式不对。`);
      if (r.block) throw blocked(String(r.reason || `隐私保护插件“${p.name}”拦截了这条回答。`).slice(0, 300), { plugin: p.name });
      if (r.answer !== undefined) out = r.answer;
    }
    return restoreText(out, session.state.byToken);
  }
  return { info, beforeSend, afterReceive, reload() { ready = null; return load(); } };
}
