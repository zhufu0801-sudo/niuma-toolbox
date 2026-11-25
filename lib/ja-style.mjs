// 办公工具箱 · 日文文体转换（2026-11-25）：です・ます体 ⇄ である体。不用 AI，用离线分词库 kuromoji（Apache-2.0）判断词性和活用，
// 只改谓语的词尾，返回每段要替换的位置。「」『』里的引用不动；“ください”这类请求句改不自然，不改，只提示。
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
let tokenizerPromise = null;
export function tokenizer() {
  if (!tokenizerPromise) tokenizerPromise = new Promise((resolve, reject) => {
    // kuromoji 已打包进 lib/vendor（不用 npm install），词典在 lib/vendor/kuromoji-dict。
    let kuromoji;
    try { kuromoji = require('./vendor/kuromoji.min.cjs'); } catch { reject(Error('缺少日文分词组件（kuromoji），请重新安装工具箱。')); return; }
    const dicPath = path.join(path.dirname(fileURLToPath(import.meta.url)), 'vendor', 'kuromoji-dict') + path.sep;
    kuromoji.builder({ dicPath }).build((err, t) => err ? reject(Error('日文分词组件加载失败：' + err.message)) : resolve(t));
  }).catch(e => { tokenizerPromise = null; throw e; });
  return tokenizerPromise;
}

// ---- 动词活用 ----
const ROW = { 'く': 'かきくけこ', 'ぐ': 'がぎぐげご', 'す': 'さしすせそ', 'つ': 'たちつてと', 'ぬ': 'なにぬねの', 'ぶ': 'ばびぶべぼ', 'む': 'まみむめも', 'る': 'らりるれろ', 'う': 'わいうえお' };
function godanTa(basic, type) {
  const stem = basic.slice(0, -1), last = basic.slice(-1);
  if (/ラ行特殊/.test(type)) return stem + 'った';
  if (/カ行促音便/.test(type)) return stem + 'った';
  return stem + ({ 'く': 'いた', 'ぐ': 'いだ', 'す': 'した', 'つ': 'った', 'ぬ': 'んだ', 'ぶ': 'んだ', 'む': 'んだ', 'る': 'った', 'う': /ウ音便/.test(type) ? 'うた' : 'った' }[last] || 'った');
}
// form: basic 原形 / renyo 连用（接ます）/ nai 否定 / ta 过去 / te 连接 / vol 意志
export function conjugate(basic, type, form) {
  type = String(type || '');
  if (/^サ変/.test(type) || /する$/.test(basic) && /サ変/.test(type)) {
    const p = basic.replace(/[すず]る$/, ''), zu = /ずる$/.test(basic);
    return p + ({ basic: zu ? 'ずる' : 'する', renyo: zu ? 'じ' : 'し', nai: zu ? 'じない' : 'しない', ta: zu ? 'じた' : 'した', te: zu ? 'じて' : 'して', vol: zu ? 'じよう' : 'しよう' }[form]);
  }
  if (/^カ変/.test(type)) { const p = basic.slice(0, -2) + (basic.slice(-2) === 'くる' ? '' : basic.slice(-2, -1)); const k = basic.slice(-2) === 'くる';
    return p + ({ basic: k ? 'くる' : 'る', renyo: k ? 'き' : '', nai: k ? 'こない' : 'ない', ta: k ? 'きた' : 'た', te: k ? 'きて' : 'て', vol: k ? 'こよう' : 'よう' }[form]); }
  if (/一段/.test(type)) { const s = basic.slice(0, -1); return s + ({ basic: 'る', renyo: '', nai: 'ない', ta: 'た', te: 'て', vol: 'よう' }[form]); }
  if (/^五段/.test(type)) {
    const stem = basic.slice(0, -1), row = ROW[basic.slice(-1)]; if (!row) return null;
    if (form === 'basic') return basic;
    if (form === 'renyo') return stem + (/ラ行特殊/.test(type) ? 'い' : row[1]);
    if (form === 'nai') return basic === 'ある' ? 'ない' : stem + row[0] + 'ない';
    if (form === 'vol') return stem + row[4] + 'う';
    const ta = godanTa(basic, type); return form === 'ta' ? ta : ta.replace(/た$/, 'て').replace(/だ$/, 'で');
  }
  return null;
}
// 谦让 / 郑重的说法换成普通说法（である体里更自然）
const PLAIN_VERB = { 'いたす': ['する', 'サ変・スル'], '致す': ['する', 'サ変・スル'], 'おる': ['いる', '一段'], '居る': ['いる', '一段'], 'ござる': ['ある', '五段・ラ行'] };

function spans(tokens, text) { let at = 0; return tokens.map(t => { const s = text.indexOf(t.surface_form, at); const start = s >= 0 ? s : at; at = start + t.surface_form.length; return { ...t, start, end: at }; }); }
function quoted(text) { const mask = new Array(text.length).fill(false); let depth = 0; for (let i = 0; i < text.length; i++) { const c = text[i]; if (c === '「' || c === '『') depth++; if (depth > 0) mask[i] = true; if ((c === '」' || c === '』') && depth > 0) depth--; } return mask; }
const isAux = (t, type) => t && t.pos === '助動詞' && (!type || t.conjugated_type === type);
const isVerbLike = t => t && (t.pos === '動詞' || t.pos === '助動詞' && /一段|五段|サ変/.test(t.conjugated_type || ''));

// ---- です・ます体 → である体 ----
function toPlain(tk, { da = false } = {}) {
  const edits = [], notes = [];
  for (let i = 0; i < tk.length; i++) {
    const t = tk[i];
    if (t.basic_form === 'くださる' && /命令/.test(t.conjugated_form || '')) { notes.push({ start: t.start, end: t.end, kind: 'request' }); continue; }
    if (isAux(t, '特殊・マス')) {
      const v = tk[i - 1]; if (!isVerbLike(v) || !/連用/.test(v.conjugated_form || '')) continue;
      let base = v.basic_form, type = v.conjugated_type; if (/^(おり|居り)$/.test(v.surface_form)) base = 'おる'; if (PLAIN_VERB[base]) [base, type] = PLAIN_VERB[base];
      const n1 = tk[i + 1], n2 = tk[i + 2], n3 = tk[i + 3]; let form = null, end = t.end, skip = 0;
      if (t.surface_form === 'ます') form = 'basic';
      else if (t.surface_form === 'まし' && n1 && n1.surface_form === 'た') { form = 'ta'; end = n1.end; skip = 1; }
      else if (t.surface_form === 'まし' && n1 && n1.surface_form === 'て') { form = 'te'; end = n1.end; skip = 1; }
      else if (t.surface_form === 'ましょ' && n1 && n1.surface_form === 'う') { form = 'vol'; end = n1.end; skip = 1; }
      else if (t.surface_form === 'ませ' && n1 && n1.surface_form === 'ん') {
        if (n2 && n2.surface_form === 'でし' && n3 && n3.surface_form === 'た') { form = 'nakatta'; end = n3.end; skip = 3; } else { form = 'nai'; end = n1.end; skip = 1; } }
      if (!form) continue;
      let out = form === 'nakatta' ? (conjugate(base, type, 'nai') || '').replace(/ない$/, 'なかった') : conjugate(base, type, form);
      if (!out) continue;
      edits.push({ start: v.start, end, text: out, from: v.surface_form + tk.slice(i, i + skip + 1).map(x => x.surface_form).join('') }); i += skip; continue;
    }
    if (isAux(t, '特殊・デス')) {
      const p = tk[i - 1], n1 = tk[i + 1]; const adjLike = p && (p.pos === '形容詞' || isAux(p, '特殊・ナイ') || isAux(p, '特殊・タイ') || isAux(p, '特殊・タ'));
      let end = t.end, out = null, skip = 0;
      if (t.surface_form === 'です') out = adjLike ? '' : (da ? 'だ' : 'である');
      else if (t.surface_form === 'でし' && n1 && n1.surface_form === 'た') { out = adjLike ? '' : (da ? 'だった' : 'であった'); end = n1.end; skip = 1; }
      else if (t.surface_form === 'でしょ' && n1 && n1.surface_form === 'う') { out = da ? 'だろう' : 'であろう'; end = n1.end; skip = 1; }
      if (out == null) continue;
      // 疑问句“〜ですか”在である体里直接写“〜か”
      if (n1 && n1.surface_form === 'か' && t.surface_form === 'です') out = '';
      edits.push({ start: t.start, end, text: out, from: tk.slice(i, i + skip + 1).map(x => x.surface_form).join('') }); i += skip;
    }
  }
  return { edits, notes };
}

// ---- である体 → です・ます体（只改句末和“が、”“ので、”等分句处的谓语）----
const FINAL_PARTICLE = new Set(['か', 'ね', 'よ', 'な', 'わ', 'ぞ']);
function predicateEnds(tk, text) {
  const ends = [];
  for (let i = 0; i < tk.length; i++) {
    const t = tk[i], next = tk[i + 1];
    const atEnd = !next || /^[。．！？!?]+$/.test(next.surface_form) || /[\r\n]/.test(text.slice(t.end, next.start));
    const beforeClause = next && next.pos === '助詞' && /接続助詞/.test(next.pos_detail_1 || '') && /^(が|けれど|けれども|ので|から|し)$/.test(next.surface_form) && tk[i + 2] && /^[、，,]$/.test(tk[i + 2].surface_form);
    if (atEnd || beforeClause) { let k = i; while (k > 0 && tk[k].pos === '助詞' && FINAL_PARTICLE.has(tk[k].surface_form) && atEnd) k--; ends.push(k); }
  }
  return ends;
}
function toPolite(tk, text) {
  const edits = [];
  for (const k of predicateEnds(tk, text)) {
    const L = tk[k], P = tk[k - 1], PP = tk[k - 2]; if (!L) continue;
    const push = (start, end, out) => edits.push({ start, end, text: out, from: text.slice(start, end) });
    if (L.surface_form === 'う' && P && /^(だろ|であろ)$/.test(P.surface_form)) { push(P.start, L.end, 'でしょう'); continue; }
    if (L.surface_form === 'う' && P && P.surface_form === 'ろ' && PP && PP.surface_form === 'であ') { push(PP.start, L.end, 'でしょう'); continue; }
    if (L.surface_form === 'う' && P && P.surface_form === 'あろ' && PP && PP.surface_form === 'で') { push(PP.start, L.end, 'でしょう'); continue; }
    if (L.pos === '名詞' && tk[k + 1] && tk[k + 1].surface_form === 'か' && tk[k + 1].pos === '助詞') { push(L.end, L.end, 'です'); continue; }
    if (isAux(L, '特殊・ダ') && L.surface_form === 'だ') { push(L.start, L.end, 'です'); continue; }
    if (L.basic_form === 'ある' && L.surface_form === 'ある' && P && P.surface_form === 'で' && (P.pos === '助動詞' || P.pos === '助詞')) { push(P.start, L.end, 'です'); continue; }
    if (isAux(L, '特殊・タ') && P) {
      if (P.surface_form === 'だっ') { push(P.start, L.end, 'でした'); continue; }
      if (P.surface_form === 'あっ' && PP && PP.surface_form === 'で') { push(PP.start, L.end, 'でした'); continue; }
      if (P.surface_form === 'なかっ') { const V = tk[k - 2];
        if ((P.pos === '助動詞') && isVerbLike(V) && /未然/.test(V.conjugated_form || '')) { const r = conjugate(V.basic_form, V.conjugated_type, 'renyo'); if (r != null) { push(V.start, L.end, r + 'ませんでした'); continue; } }
        if (P.pos === '助動詞' && V && V.pos === '形容詞') { push(P.start, L.end, 'ありませんでした'); continue; }
        if (P.pos === '形容詞') { if (V && V.surface_form === 'は' && tk[k - 3] && tk[k - 3].surface_form === 'で') push(tk[k - 3].start, L.end, 'ではありませんでした'); else push(P.start, L.end, 'ありませんでした'); continue; } }
      if (P.pos === '形容詞' || isAux(P, '特殊・タイ')) { push(L.end, L.end, 'です'); continue; }
      if (isVerbLike(P)) { const r = conjugate(P.basic_form, P.conjugated_type, 'renyo'); if (r != null) { push(P.start, L.end, r + 'ました'); continue; } }
      continue;
    }
    if (L.surface_form === 'ない' && (isAux(L, '特殊・ナイ') || L.pos === '形容詞')) {
      if (isAux(L, '特殊・ナイ') && isVerbLike(P) && /未然/.test(P.conjugated_form || '')) { const r = conjugate(P.basic_form, P.conjugated_type, 'renyo'); if (r != null) { push(P.start, L.end, r + 'ません'); continue; } }
      if (P && P.surface_form === 'は' && PP && PP.surface_form === 'で') { push(PP.start, L.end, 'ではありません'); continue; }
      if (P && P.surface_form === 'で') { push(P.start, L.end, 'ではありません'); continue; }
      if (!isVerbLike(P)) { push(L.start, L.end, 'ありません'); continue; }
      continue;
    }
    if ((L.surface_form === 'う' || L.surface_form === 'よう') && L.pos === '助動詞' && isVerbLike(P) && /未然ウ接続|未然/.test(P.conjugated_form || '')) { const r = conjugate(P.basic_form, P.conjugated_type, 'renyo'); if (r != null) { push(P.start, L.end, r + 'ましょう'); continue; } }
    if (L.pos === '形容詞' && L.conjugated_form === '基本形' || isAux(L, '特殊・タイ') && L.conjugated_form === '基本形' || L.pos === '助動詞' && L.basic_form === 'らしい') { push(L.end, L.end, 'です'); continue; }
    if (isVerbLike(L) && L.conjugated_form === '基本形') { if (L.basic_form === 'ある' && P && P.surface_form === 'で') { push(P.start, L.end, 'です'); continue; } const r = conjugate(L.basic_form, L.conjugated_type, 'renyo'); if (r != null) { push(L.start, L.end, r + 'ます'); continue; } }
  }
  return { edits, notes: [] };
}

export async function convertParagraphs({ paragraphs, to, plain }) {
  const list = [].concat(paragraphs || []).map(x => String(x == null ? '' : x));
  if (!list.length) throw Error('没有要转换的文字。');
  if (list.length > 5000 || list.reduce((s, x) => s + x.length, 0) > 400000) throw Error('文字太多，请选中一部分再转换。');
  const tk = await tokenizer(), out = [];
  for (const text of list) {
    if (!/[぀-ヿ]/.test(text)) { out.push({ edits: [], notes: [] }); continue; }
    const tokens = spans(tk.tokenize(text), text), mask = quoted(text);
    const r = to === 'polite' ? toPolite(tokens, text) : toPlain(tokens, { da: plain === 'da' });
    const keep = e => !mask.slice(e.start, Math.max(e.end, e.start + 1)).some(Boolean);
    out.push({ edits: r.edits.filter(keep).sort((a, b) => a.start - b.start), notes: r.notes.filter(keep) });
  }
  return { paragraphs: out, edits: out.reduce((s, p) => s + p.edits.length, 0), notes: out.reduce((s, p) => s + p.notes.length, 0) };
}
export function applyEdits(text, edits) { let s = String(text); [...edits].sort((a, b) => b.start - a.start).forEach(e => { s = s.slice(0, e.start) + e.text + s.slice(e.end); }); return s; }
