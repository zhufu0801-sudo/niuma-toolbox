// 住所 ↔ 郵便番号（2026-11-25）：用日本邮政的邮编数据（utf_ken_all.zip）。
// 日本邮政写明：郵便番号データに限っては著作権を主張しません。自由に配布していただいて結構です。
// 数据直接放进安装包（lib/data/postal-jp.json.gz）。每月更新，可以点“更新数据”从官网重新下载，
// 或导入自己下载的 zip（存在 generated/，用两份里较新的）。只到町域（丁目、番地不在官方数据里）。
import fs from 'node:fs/promises';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';

// 2026-11-25：日本邮政改了网址，旧地址（/zipcode/dl/utf/zip/）已经 404。以后再改的话，会去下载页面上找 utf_ken_all.zip 的链接。
export const OFFICIAL_URL = 'https://www.post.japanpost.jp/service/search/zipcode/download/utf/zip/utf_ken_all.zip';
export const OFFICIAL_PAGE = 'https://www.post.japanpost.jp/service/search/zipcode/download/utf-zip.html';
const OLD_URLS = ['https://www.post.japanpost.jp/zipcode/dl/utf/zip/utf_ken_all.zip'];
const HERE = path.dirname(fileURLToPath(import.meta.url));
export const BUNDLED = path.join(HERE, 'data', 'postal-jp.json.gz');
export const PREFS = ['北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県', '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県', '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県', '岐阜県', '静岡県', '愛知県', '三重県', '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県', '鳥取県', '島根県', '岡山県', '広島県', '山口県', '徳島県', '香川県', '愛媛県', '高知県', '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県', '沖縄県'];

// ---------- 公式 CSV → 小さいデータ ----------
// zip（1 ファイルだけ入っているもの）から中身を取り出す。外部ライブラリは使わない。
export function unzipFirst(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) throw Error('读不出这个 zip 文件。');
  const cdOffset = buf.readUInt32LE(eocd + 16), entries = buf.readUInt16LE(eocd + 10);
  let p = cdOffset;
  for (let n = 0; n < entries; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw Error('zip 文件已损坏。');
    const method = buf.readUInt16LE(p + 10), dos = buf.readUInt16LE(p + 14), size = buf.readUInt32LE(p + 20), nameLen = buf.readUInt16LE(p + 28), extraLen = buf.readUInt16LE(p + 30), commentLen = buf.readUInt16LE(p + 32), local = buf.readUInt32LE(p + 42);
    const name = buf.subarray(p + 46, p + 46 + nameLen).toString('utf8');
    p += 46 + nameLen + extraLen + commentLen;
    if (!/\.csv$/i.test(name)) continue;
    const lNameLen = buf.readUInt16LE(local + 26), lExtra = buf.readUInt16LE(local + 28), start = local + 30 + lNameLen + lExtra, data = buf.subarray(start, start + size);
    const date = dosDate(dos);
    if (method === 0) return { name, date, data: Buffer.from(data) };
    if (method === 8) return { name, date, data: zlib.inflateRawSync(data) };
    throw Error('不支持这种压缩方式（' + method + '）。');
  }
  throw Error('zip 里没有 CSV 文件。');
}
// zip 里文件的日期（日本邮政打包那天）→ 2026-10-30；没有或不像真的日期就返回 ''。
function dosDate(d) { const y = 1980 + (d >> 9), m = (d >> 5) & 15, day = d & 31; if (y < 2000 || m < 1 || m > 12 || day < 1) return ''; return y + '-' + String(m).padStart(2, '0') + '-' + String(day).padStart(2, '0'); }
const today = () => new Date().toISOString().slice(0, 10);
// 数据日期：调用方指定 > zip 里的文件日期 > 网站给的日期 > 今天
function withDate(buf, { date = '', fallback = '' } = {}) { const z = unzipFirst(buf); return { text: z.data.toString('utf8'), date: date || z.date || fallback || today() }; }
function csvRows(text) {
  const rows = []; let row = [], field = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; } else field += c; continue; }
    if (c === '"') q = true; else if (c === ',') { row.push(field); field = ''; } else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(field); if (row.length > 1) rows.push(row); row = []; field = ''; } else field += c;
  }
  if (field || row.length) { row.push(field); if (row.length > 1) rows.push(row); }
  return rows;
}
// 町域名の後始末：「以下に掲載がない場合」「○○の次に番地がくる場合」「○○一円」は町域なし；（ ）の中は補足として分ける。
export function cleanTown(town, city) {
  let t = String(town || '').trim();
  const whole = /一円$/.test(t) && (t === city + '一円' || t.replace(/一円$/, '') === city.replace(/^.*?郡/, ''));
  if (!t || t === '以下に掲載がない場合' || /の次に番地がくる場合$/.test(t) || whole) return { town: '', note: '' };
  const m = /^([^（(]*)[（(](.*)[）)]$/.exec(t);
  if (m) { const base = m[1].trim(), note = m[2].trim(); if (/^次のビルを除く$|^その他$|^次を除く/.test(note)) return { town: base, note: '' }; const building = /階|地階/.test(note); return { town: base, note, building }; }
  return { town: t, note: '' };
}
// 公式 CSV（utf_ken_all.csv、1 レコード 1 行、UTF-8）→ { prefs, cities:[[都道府県番号, 市区町村名]], rows:[[郵便番号(数値), 市区町村番号, 町域, 補足, ビル]] }
export function buildFromCsv(text, { date = '' } = {}) {
  const cities = [], cityIndex = new Map(), rows = [], seen = new Set();
  for (const r of csvRows(String(text).replace(/^﻿/, ''))) {
    if (r.length < 9 || !/^\d{7}$/.test(r[2])) continue;
    const pref = PREFS.indexOf(r[6]); if (pref < 0) continue;
    const key = pref + '\u0001' + r[7];
    let ci = cityIndex.get(key); if (ci === undefined) { ci = cities.length; cities.push([pref, r[7]]); cityIndex.set(key, ci); }
    const t = cleanTown(r[8], r[7]), row = [Number(r[2]), ci, t.town];
    if (t.note || t.building) row.push(t.note || '');
    if (t.building) row.push(1);
    const k = row.join('\u0001'); if (seen.has(k)) continue; seen.add(k);
    rows.push(row);
  }
  if (rows.length < 100) throw Error('这不是日本邮政的邮编数据（只有 ' + rows.length + ' 行）。请用 utf_ken_all.zip。');
  return { v: 1, source: OFFICIAL_URL, date, prefs: PREFS, cities, rows };
}
export async function writeData(data, file) { await fs.mkdir(path.dirname(file), { recursive: true }); const tmp = file + '.tmp'; await fs.writeFile(tmp, zlib.gzipSync(Buffer.from(JSON.stringify(data)), { level: 9 })); await fs.rename(tmp, file); }
export async function dataFromFile(file, opts = {}) {
  const buf = await fs.readFile(file);
  let text, date;
  if (buf.length >= 4 && buf.readUInt32LE(0) === 0x04034b50) ({ text, date } = withDate(buf, opts));
  else { text = buf.toString('utf8'); date = opts.date || (await fs.stat(file)).mtime.toISOString().slice(0, 10); }
  if (/[�]/.test(text.slice(0, 2000))) throw Error('文字编码不对，请用 UTF-8 版的 utf_ken_all.zip。');
  return buildFromCsv(text, { date });
}
// 公式サイトからダウンロード（利用者のパソコン上で動く）。
// 顺序：现在的网址 → 下载页面上找到的链接（网址又变了的时候）→ 旧网址。
export function findZipLink(html, base = OFFICIAL_PAGE) {
  const m = /href\s*=\s*["']([^"']*utf_ken_all\.zip)["']/i.exec(String(html || ''));
  if (!m) return '';
  try { const u = new URL(m[1], base); return u.protocol === 'https:' && /(^|\.)post\.japanpost\.jp$/.test(u.hostname) ? u.href : ''; } catch { return ''; }
}
export async function downloadOfficial({ fetchImpl = fetch, timeout = 60000 } = {}) {
  const ctl = new AbortController(), timer = setTimeout(() => ctl.abort(), timeout), tried = [];
  async function get(url) { const r = await fetchImpl(url, { signal: ctl.signal }); if (!r.ok) throw Error('HTTP ' + r.status); return r; }
  async function zipFrom(url) {
    const r = await get(url), buf = Buffer.from(await r.arrayBuffer());
    if (buf.length < 4 || buf.readUInt32LE(0) !== 0x04034b50) throw Error('下载到的不是 zip');
    const lm = r.headers.get('last-modified'), fallback = lm && !isNaN(new Date(lm)) ? new Date(lm).toISOString().slice(0, 10) : '';
    const { text, date } = withDate(buf, { fallback });
    return buildFromCsv(text, { date });
  }
  try {
    try { return await zipFrom(OFFICIAL_URL); } catch (e) { if (e.name === 'AbortError') throw e; tried.push(e.message); }
    try { const link = findZipLink(await (await get(OFFICIAL_PAGE)).text()); if (link && link !== OFFICIAL_URL) return await zipFrom(link); tried.push('下载页面上没找到 utf_ken_all.zip'); } catch (e) { if (e.name === 'AbortError') throw e; tried.push(e.message); }
    for (const u of OLD_URLS) { try { return await zipFrom(u); } catch (e) { if (e.name === 'AbortError') throw e; tried.push(e.message); } }
    throw Error([...new Set(tried)].join('；'));
  } catch (e) { throw Error('没能从日本邮政网站下载（' + (e.name === 'AbortError' ? '超时' : e.message) + '）。请检查网络，或者在 ' + OFFICIAL_PAGE + ' 下载 utf_ken_all.zip 后导入。'); }
  finally { clearTimeout(timer); }
}

// ---------- 検索 ----------
// 表記ゆれ：全角英数→半角（NFKC）、空白・〒・番号の区切りを除く、ヶ/ケ/が（地名の「ヶ」）を同じに。
export function norm(s) { return String(s || '').normalize('NFKC').replace(/[\s　]/g, '').replace(/[ヶヵケ]/g, 'ケ').replace(/之/g, 'の').replace(/[‐－―ー−]/g, '-'); }
const zip7 = n => String(n).padStart(7, '0');
const KAN = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
function kanji(s) { if (/^\d+$/.test(s)) return Number(s); let n = 0, cur = 0; for (const c of s) { if (c === '十') { n += (cur || 1) * 10; cur = 0; } else if (KAN[c]) cur = KAN[c]; else return 0; } return n + cur; }
// 住所の町域の後ろから丁目の数字を取る：「5丁目」「五丁目」「2-1」「2番」
function chomeOf(rest) { const m = /^(\d+|[一二三四五六七八九十]+)(丁目|丁|-|番|$)/.exec(rest); return m ? kanji(m[1]) : 0; }
// 補足（「１〜１９丁目」「１、３丁目」「２丁目」）がその丁目を含むか
function covers(note, n) { const t = String(note || '').normalize('NFKC'); if (!/丁目/.test(t)) return false; for (const part of t.split(/[、,・]/)) { const r = /(\d+)(?:丁目)?\s*[〜~～-]\s*(\d+)丁目/.exec(part); if (r && n >= Number(r[1]) && n <= Number(r[2])) return true; const one = /(\d+)丁目/.exec(part) || /^(\d+)$/.exec(part.trim()); if (one && Number(one[1]) === n) return true; } return false; }
export const formatZip = z => { const s = zip7(z); return s.slice(0, 3) + '-' + s.slice(3); };
export function createPostal(data) {
  const prefs = data.prefs, cities = data.cities, rows = data.rows;
  const byZip = new Map(), byCity = new Map();
  rows.forEach((r, i) => { (byZip.get(r[0]) || byZip.set(r[0], []).get(r[0])).push(i); (byCity.get(r[1]) || byCity.set(r[1], []).get(r[1])).push(i); });
  // 郡名・島名を省いた書き方も認める（西多摩郡日の出町 → 日の出町、三宅島三宅村 → 三宅村）
  const cityNorm = cities.map(c => ({ full: norm(c[1]), short: norm(c[1].replace(/^.+?郡/, '').replace(/^.+?島(?=.+[町村]$)/, '')) }));
  const prefNorm = prefs.map(norm);
  const out = i => { const r = rows[i], c = cities[r[1]]; return { zip: formatZip(r[0]), pref: prefs[c[0]], city: c[1], town: r[2], note: r[3] || '', building: !!r[4], address: prefs[c[0]] + c[1] + r[2] }; };
  function byZipCode(input) {
    const d = String(input || '').normalize('NFKC').replace(/[^\d]/g, '');
    if (d.length !== 7) return { ok: false, reason: 'format', candidates: [] };
    const list = (byZip.get(Number(d)) || []).map(out);
    return { ok: list.length > 0, reason: list.length ? '' : 'notfound', candidates: list };
  }
  function byAddress(input) {
    let s = norm(input).replace(/^〒?\d{3}-?\d{4}/, '');
    if (!s) return { ok: false, reason: 'empty', candidates: [] };
    let pi = prefNorm.findIndex(p => s.startsWith(p));
    if (pi < 0) { const short = prefNorm.map(p => p.replace(/[都府県]$/, '')).findIndex(p => p.length > 1 && s.startsWith(p) && !s.startsWith(p + '市')); if (short >= 0 && !/^(京都|大阪)/.test(s)) pi = short; }
    if (pi >= 0) s = s.slice(s.startsWith(prefNorm[pi]) ? prefNorm[pi].length : prefNorm[pi].replace(/[都府県]$/, '').length);
    // 市区町村：都道府県が分かればその中で、分からなければ全国で、一番長く一致するもの（郡名の省略も可）
    let best = [], bestLen = 0;
    cities.forEach((c, ci) => { if (pi >= 0 && c[0] !== pi) return; for (const k of new Set([cityNorm[ci].full, cityNorm[ci].short])) if (k && s.startsWith(k)) { if (k.length > bestLen) { best = [{ ci, len: k.length }]; bestLen = k.length; } else if (k.length === bestLen && !best.some(b => b.ci === ci)) best.push({ ci, len: k.length }); } });
    if (!best.length) return { ok: false, reason: 'city', candidates: [] };
    const found = [];
    for (const b of best) {
      const rest = s.slice(b.len), idx = byCity.get(b.ci) || [];
      let tLen = -1, hits = [];
      for (const i of idx) { const t = norm(rows[i][2]); if (!t || rows[i][4]) continue; if (rest.startsWith(t)) { if (t.length > tLen) { tLen = t.length; hits = [i]; } else if (t.length === tLen) hits.push(i); } }
      // 同じ町域が丁目で分かれているとき（大通西（１〜１９丁目）など）、住所の丁目の数字で絞る
      if (hits.length > 1) { const n = chomeOf(rest.slice(tLen)); if (n) { const fit = hits.filter(i => covers(rows[i][3], n)); if (fit.length) hits = fit; } }
      if (hits.length) found.push(...hits.map(i => ({ i, level: 'town' })));
      else { const cityRow = idx.filter(i => !rows[i][2]); if (cityRow.length) found.push(...cityRow.map(i => ({ i, level: 'city' }))); else found.push(...[...new Set(idx.map(i => rows[i][0]))].length === 1 ? [{ i: idx[0], level: 'city' }] : []); }
    }
    const uniq = [], seen = new Set(); for (const f of found) { const o = out(f.i), k = o.zip + o.address + o.note; if (!seen.has(k)) { seen.add(k); uniq.push({ ...o, level: f.level }); } }
    return { ok: uniq.length > 0, reason: uniq.length ? (best.length > 1 && pi < 0 ? 'ambiguous-city' : '') : 'town', candidates: uniq, exact: uniq.length === 1 };
  }
  return { byZip: byZipCode, byAddress, info: { date: data.date || '', rows: rows.length } };
}

// 本机服务用：同梱データと generated/ の取り込みデータのうち新しい方を読む。
export function createPostalStore({ userFile, bundled = BUNDLED } = {}) {
  let cache = null, loading = null;
  async function readGz(f) { try { return JSON.parse(zlib.gunzipSync(await fs.readFile(f)).toString('utf8')); } catch (e) { if (e.code === 'ENOENT') return null; throw Error('邮编数据文件损坏（' + path.basename(f) + '），请重新更新数据。'); } }
  async function load() {
    if (cache) return cache;
    if (!loading) loading = (async () => {
      const [a, b] = await Promise.all([userFile ? readGz(userFile) : null, readGz(bundled)]);
      const data = a && (!b || String(a.date) >= String(b.date)) ? a : b;
      if (!data) throw Error('没有邮编数据。请点“更新数据”从日本邮政网站下载，或者导入 utf_ken_all.zip。');
      cache = { postal: createPostal(data), from: data === a ? 'user' : 'bundled' }; return cache;
    })().finally(() => { loading = null; });
    return loading;
  }
  async function replace(data) { if (!userFile) throw Error('没有保存位置。'); await writeData(data, userFile); cache = null; return (await load()).postal.info; }
  return {
    async status() { try { const c = await load(); return { ok: true, ...c.postal.info, from: c.from }; } catch (e) { return { ok: false, error: e.message }; } },
    async zip(list) { const p = (await load()).postal; return list.map(x => p.byZip(x)); },
    async address(list) { const p = (await load()).postal; return list.map(x => p.byAddress(x)); },
    async update(opts) { return replace(await downloadOfficial(opts)); },
    async importFile(file) { return replace(await dataFromFile(file)); }
  };
}
