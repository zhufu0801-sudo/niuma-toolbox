// 粘贴栏：把复制过、收集过的文字和表格排成一列，点一下放进当前文档。
// 只保存在本机 generated/pastebar.json。普通条目最多 30 条（新的在前，旧的自动挤掉），钉住的另算、最多 20 条。
// 自动记录复制内容默认关闭；打开后也会跳过像密码、身份证号、银行卡号、マイナンバー、API Key 的内容。
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { checkCardBlock, cardTitle } from './cards.mjs';

export const PASTEBAR_LIMIT = 30;
export const PASTEBAR_PIN_LIMIT = 20;
const SOURCES = ['wps', 'et', 'wpp', 'clipboard'];

function luhn(digits) {
  let sum = 0, double = false;
  for (let i = digits.length - 1; i >= 0; i--) { let d = Number(digits[i]); if (double) { d *= 2; if (d > 9) d -= 9; } sum += d; double = !double; }
  return sum % 10 === 0;
}

// 自动记录时跳过的内容。只看“像不像”，宁可多跳过；手动收集不受影响（客户自己决定）。
export function sensitiveReason(text) {
  const s = String(text ?? '').trim();
  if (!s) return 'empty';
  if (/(^|[^\d])\d{6}(19|20)\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\d{3}[\dXx]($|[^\d])/.test(s)) return 'id';
  for (const m of s.matchAll(/(?:\d[ -]?){12,18}\d/g)) { const d = m[0].replace(/\D/g, ''); if (d.length >= 13 && d.length <= 19 && luhn(d)) return 'card'; }
  if (/^\d{4}[ -]?\d{4}[ -]?\d{4}$/.test(s)) return 'mynumber';
  if (/^(sk|pk|rk)-[A-Za-z0-9_-]{16,}$|^(ghp|gho|ghs|github_pat)_[A-Za-z0-9_]{16,}$|^xox[abpr]-[A-Za-z0-9-]{10,}$|^AKIA[0-9A-Z]{16}$|^AIza[0-9A-Za-z_-]{30,}$/.test(s)) return 'key';
  if (!/\s/.test(s) && s.length >= 32 && /^[A-Za-z0-9+/=_-]+$/.test(s) && /\d/.test(s) && /[A-Za-z]/.test(s)) return 'key';
  // 一行、没有空格和中日文、8–64 个字符、同时有字母和数字（再加大小写或符号）的，像密码。
  if (!/\s/.test(s) && s.length >= 8 && s.length <= 64 && !/[぀-ヿ㐀-鿿]/.test(s) && !/^https?:\/\//i.test(s) && !/^[\w.+-]+@[\w-]+(\.[\w-]+)+$/.test(s)) {
    const kinds = [/[a-z]/, /[A-Z]/, /\d/, /[^A-Za-z0-9]/].filter(r => r.test(s)).length;
    if (/[A-Za-z]/.test(s) && /\d/.test(s) && kinds >= 3) return 'password';
  }
  return null;
}

// 从表格复制的文字是“制表符分列、换行分行”，整齐的就存成表格，插进文字文档时是一张真表格。
export function blockFromText(text) {
  const s = String(text ?? '').replace(/\r\n?/g, '\n').replace(/\n+$/, '');
  const lines = s.split('\n');
  if (lines.length >= 1 && lines.every(l => l.includes('\t'))) {
    const rows = lines.map(l => l.split('\t')), width = rows[0].length;
    if (width >= 2 && width <= 10 && rows.length <= 20 && rows.every(r => r.length === width)) return { type: 'table', rows };
  }
  return { type: 'text', text: s };
}

function sameBlock(a, b) { return JSON.stringify(a) === JSON.stringify(b); }

export function createPasteBar(file, { limit = PASTEBAR_LIMIT, pinLimit = PASTEBAR_PIN_LIMIT, now = () => Date.now() } = {}) {
  let items = null, queue = Promise.resolve();
  const serial = fn => { const result = queue.catch(() => {}).then(fn); queue = result; return result; };
  async function load() {
    if (items) return items;
    if (!file) return (items = []);
    try { const data = JSON.parse(await fs.readFile(file, 'utf8')); if (!Array.isArray(data.items)) throw Error(); items = data.items; }
    catch (e) { if (e.code === 'ENOENT') items = []; else throw Error('粘贴栏记录无法读取，已保留原文件。'); }
    return items;
  }
  async function save(next) {
    items = next;
    if (!file) return;
    await fs.mkdir(path.dirname(file), { recursive: true });
    const temp = file + '.' + randomBytes(4).toString('hex') + '.tmp';
    await fs.writeFile(temp, JSON.stringify({ items: next }), 'utf8');
    await fs.rename(temp, file);
  }
  function trim(list) {
    let plain = 0;
    return list.filter(x => x.pinned || ++plain <= limit);
  }
  return {
    list() { return serial(async () => (await load()).slice()); },
    // from：manual（侧栏“收集选中内容”）或 clipboard（复制时自动记录 / 从剪贴板添加）。
    add(input, sourceHost, from = 'manual') {
      return serial(async () => {
        if (!SOURCES.includes(sourceHost)) throw Error('来源无效。');
        let block = input && input.type ? input : blockFromText(input);
        if (block.type === 'text' && block.text.length > 10000) throw Error('粘贴栏每条最多 1 万字。');
        block = checkCardBlock(block);
        const list = (await load()).slice(), old = list.findIndex(x => sameBlock(x.block, block));
        let item;
        if (old >= 0) { item = { ...list[old], at: now() }; list.splice(old, 1); }
        else item = { id: now().toString(36) + randomBytes(3).toString('hex'), block, title: cardTitle(block), sourceHost, from: from === 'clipboard' ? 'clipboard' : 'manual', pinned: false, at: now() };
        // 钉住的排在最前，其余按时间。
        const pins = list.filter(x => x.pinned), rest = list.filter(x => !x.pinned);
        const next = item.pinned ? [item, ...pins, ...rest] : [...pins, item, ...rest];
        await save(trim(next));
        return item;
      });
    },
    pin(id, pinned) {
      return serial(async () => {
        const list = (await load()).slice(), i = list.findIndex(x => x.id === id);
        if (i < 0) throw Error('这条内容已经不在粘贴栏里了。');
        if (pinned && !list[i].pinned && list.filter(x => x.pinned).length >= pinLimit) throw Error(`最多钉住 ${pinLimit} 条，请先取消一些。`);
        const item = { ...list[i], pinned: !!pinned }; list.splice(i, 1);
        const pins = list.filter(x => x.pinned), rest = list.filter(x => !x.pinned);
        await save(trim(item.pinned ? [item, ...pins, ...rest] : [...pins, item, ...rest]));
        return item;
      });
    },
    remove(id) {
      return serial(async () => { const list = await load(), next = list.filter(x => x.id !== id); if (next.length === list.length) return false; await save(next); return true; });
    },
    // 数据脱敏（2026-11-25）：把一条的内容换成脱敏后的（或还原回原文），位置和钉住状态不变。
    replaceBlock(id, input) {
      return serial(async () => {
        const list = (await load()).slice(), i = list.findIndex(x => x.id === id);
        if (i < 0) throw Error('这条内容已经不在粘贴栏里了。');
        const block = checkCardBlock(input);
        list[i] = { ...list[i], block, title: cardTitle(block) };
        await save(list);
        return list[i];
      });
    },
    // 清空时保留钉住的。
    clear() { return serial(async () => { const next = (await load()).filter(x => x.pinned); await save(next); return next.length; }); }
  };
}
