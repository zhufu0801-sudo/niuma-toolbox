// AI 任务：翻译润色、帮写公式。提示词、结果清理和安全检查都在这里，方便测试。
// 所有请求都经 ai.chat 发送，因此同样经过隐私保护关卡。

export const TRANSLATE_ACTIONS = {
  auto: { zh: '自动互译（中文→日文，日文→中文）', ja: '自動翻訳（中国語⇄日本語）' },
  zh2ja: { zh: '翻译成日文', ja: '日本語に翻訳' },
  ja2zh: { zh: '翻译成中文', ja: '中国語に翻訳' },
  keigo: { zh: '改成日文商务敬语', ja: 'ビジネス敬語に整える' },
  polishZh: { zh: '中文润色', ja: '中国語を推敲' },
  polishJa: { zh: '日文润色', ja: '日本語を推敲' },
  en: { zh: '翻译成英文', ja: '英語に翻訳' }
};
const TASK = {
  auto: '如果原文主要是中文，翻译成自然的日文商务表达；如果原文主要是日文，翻译成自然的简体中文。',
  zh2ja: '把原文翻译成日文。用日本企业日常商务文书的自然表达，默认用です・ます体。',
  ja2zh: '把原文翻译成简体中文，用中国企业日常商务文书的自然表达。',
  keigo: '把原文改写成得体的日文商务敬语（适当使用尊敬語・謙譲語・丁寧語），保持原意，不要过度客套。原文如果是中文，先译成日文再改成敬语。',
  polishZh: '润色这段简体中文：更通顺、准确、正式，保持原意和事实，不增加新内容。',
  polishJa: 'この日本語を推敲してください：自然で正確なビジネス文に整え、意味や事実は変えず、内容を追加しないこと。',
  en: '把原文翻译成自然、专业的英文商务表达。'
};
export function translatePrompt({ text, action, lang }) {
  const task = TASK[action] || TASK.auto;
  const sys = '你是专业的中日商务翻译和文字编辑。只输出处理后的文字本身：不要解释、不要加标题、引号或“译文：”之类的前缀。'
    + '保留原文的换行和段落结构；数字、金额、日期、产品名、人名保持准确；形如 [PHONE_1] 的方括号占位符原样保留。'
    + (lang === 'ja' ? 'ユーザーは日本語話者です。' : '');
  return [{ role: 'system', content: sys }, { role: 'user', content: task + '\n\n原文：\n' + text }];
}
// 表格：多个单元格一次处理，要求按同样顺序返回同样数量的 JSON 数组。
export function translateItemsPrompt({ items, action, lang }) {
  const task = TASK[action] || TASK.auto;
  const sys = '你是专业的中日商务翻译和文字编辑。输入是一个 JSON 字符串数组，每个元素是表格里的一个单元格。'
    + '逐个处理后，只输出同样长度、同样顺序的 JSON 字符串数组，不要解释。数字、型号、人名保持准确；形如 [PHONE_1] 的占位符原样保留。'
    + (lang === 'ja' ? 'ユーザーは日本語話者です。' : '');
  return [{ role: 'system', content: sys }, { role: 'user', content: task + '\n\n' + JSON.stringify(items) }];
}
export function parseItems(answer, count) {
  const s = String(answer || ''), a = s.indexOf('['), b = s.lastIndexOf(']');
  let list = null;
  if (a >= 0 && b > a) { try { list = JSON.parse(s.slice(a, b + 1)); } catch {} }
  if (!Array.isArray(list) || list.length !== count || list.some(x => typeof x !== 'string' && typeof x !== 'number'))
    throw Object.assign(Error('AI 返回的条数和选中的单元格对不上，请少选一些单元格再试，或换一个模型。'), { status: 502 });
  return list.map(x => String(x).trim());
}
export function cleanTranslation(answer, original) {
  let s = String(answer || '').trim();
  s = s.replace(/^```[a-zA-Z]*\n?/, '').replace(/\n?```$/, '').trim();
  s = s.replace(/^(译文|翻译|结果|润色后|修改后|訳文|翻訳|修正後|推敲後)\s*[:：]\s*/u, '');
  if (/^[「“"『].*[」”"』]$/su.test(s) && !/^[「“"『]/u.test(String(original).trim())) s = s.slice(1, -1).trim();
  if (!s) throw Object.assign(Error('AI 没有返回结果，请重试或换一个模型。'), { status: 502 });
  return s;
}

// ---------- 写作：总结、续写、扩写、精简、整理成条目、改成书面语 ----------
export const WRITE_ACTIONS = {
  summary: { zh: '总结要点', ja: '要点をまとめる' },
  continue: { zh: '续写', ja: '続きを書く' },
  expand: { zh: '扩写', ja: '詳しく書き足す' },
  shorten: { zh: '精简改短', ja: '短く簡潔に' },
  bullets: { zh: '整理成条目', ja: '箇条書きに整理' },
  formal: { zh: '改成正式书面语', ja: '丁寧な書き言葉に' }
};
const WRITE_TASK = {
  summary: '总结这段内容：先用一句话概括，再列出 3～7 条要点（每条一行，以“- ”开头）。只根据原文，不添加原文没有的信息。',
  continue: '接着原文最后一句往下写一到三段，保持主题、人称、文风和语气一致。只输出续写的新内容，不要重复原文。',
  expand: '把这段内容扩写得更充实：补充必要的说明、例子或过渡，篇幅约为原文的 1.5～2 倍。保持原意和事实，不要编造具体的数字、日期、人名或引用。输出扩写后的完整文字。',
  shorten: '把这段内容精简到原文一半左右的长度，保留关键信息和结论。',
  bullets: '把这段内容整理成清晰的条目（每条一行，以“- ”开头），需要时可以分几个小标题。不添加原文没有的信息。',
  formal: '把这段内容改写成正式、得体的书面语（日文用です・ます体），保持原意和事实。'
};
export function isWriteAction(action) { return Object.prototype.hasOwnProperty.call(WRITE_TASK, action); }
export function writePrompt({ text, action, lang, extra = '' }) {
  const task = WRITE_TASK[action];
  if (!task) throw Object.assign(Error('处理方式不对，请重新选择。'), { status: 400 });
  const sys = '你是专业的中日商务写作助手。用原文的语言输出（中文用简体，日文用日文）。只输出结果文字本身：不要解释，不要加“总结：”“续写：”之类的前缀，不要用 Markdown 的 # 标题和 ** 加粗。'
    + '数字、金额、日期、人名保持准确；形如 [PHONE_1] 的方括号占位符原样保留。'
    + (lang === 'ja' ? 'ユーザーは日本語話者です。' : '');
  const more = String(extra || '').trim().slice(0, 300);
  return [{ role: 'system', content: sys }, { role: 'user', content: task + (more ? '\n补充要求：' + more : '') + '\n\n原文：\n' + text }];
}
export function cleanWrite(answer) {
  let s = String(answer || '').trim();
  s = s.replace(/^```[a-zA-Z]*\n?/, '').replace(/\n?```$/, '').trim();
  s = s.replace(/^(总结|摘要|要点|续写|扩写|精简|改写|结果|正文|まとめ|要約|要点|続き|続編|加筆|短縮|書き換え|結果)(后的内容|内容)?\s*[:：]\s*\n?/u, '');
  s = s.replace(/\*\*([^*\n]+)\*\*/g, '$1').replace(/^#{1,6}\s+/gm, '');
  if (!s) throw Object.assign(Error('AI 没有返回结果，请重试或换一个模型。'), { status: 502 });
  return s;
}

// ---------- AI 对话 ----------
// 只保留最近的对话，单条 4000 字以内、合计 20000 字以内；可以附带用户选中的内容（8000 字以内）。
export const CHAT_LIMITS = { turns: 30, each: 4000, total: 20000, context: 8000 };
export function chatPrompt({ messages, lang, context = '' }) {
  if (!Array.isArray(messages) || !messages.length) throw Object.assign(Error('请先输入问题。'), { status: 400 });
  let list = messages.slice(-CHAT_LIMITS.turns).filter(m => m && (m.role === 'user' || m.role === 'assistant')).map(m => ({ role: m.role, content: String(m.content || '').trim() })).filter(m => m.content);
  if (!list.length || list[list.length - 1].role !== 'user') throw Object.assign(Error('请先输入问题。'), { status: 400 });
  if (list[list.length - 1].content.length > CHAT_LIMITS.each) throw Object.assign(Error('问题太长了，请控制在 4000 字以内；长文可以先选中，再勾选“附带选中的内容”。'), { status: 400 });
  list = list.map(m => m.content.length > CHAT_LIMITS.each ? { role: m.role, content: m.content.slice(0, CHAT_LIMITS.each) + '…' } : m);
  let total = 0, keep = [];
  for (let i = list.length - 1; i >= 0; i--) { total += list[i].content.length; if (total > CHAT_LIMITS.total && keep.length) break; keep.unshift(list[i]); }
  while (keep.length && keep[0].role !== 'user') keep.shift();
  const ctx = String(context || '').trim();
  if (ctx.length > CHAT_LIMITS.context) throw Object.assign(Error('附带的选中内容太长了（超过 8000 字），请少选一些。'), { status: 400 });
  if (ctx) { const last = keep[keep.length - 1]; keep[keep.length - 1] = { role: 'user', content: last.content + '\n\n——以下是我在文档里选中的内容——\n' + ctx }; }
  const sys = '你是“办公工具箱”里的 AI 助手，帮助用户在 WPS 和 Microsoft Office 里写文档、做表格、做演示：解答操作方法、写公式、起草和修改文字、翻译中日文。'
    + '回答要准确、简洁、可以直接使用；操作步骤用编号列出；公式单独一行。用户用什么语言提问就用什么语言回答（中文用简体）。'
    + '不知道或不确定的要直说，不要编造菜单名称、函数或事实。形如 [PHONE_1] 的方括号占位符原样保留。'
    + (lang === 'ja' ? 'ユーザーは日本語話者です。' : '');
  return [{ role: 'system', content: sys }, ...keep];
}

// ---------- 帮写公式 ----------
// 发给 AI 的只有：表头名、所在列、少量示例行（最多 5 行）、选区地址。不发整张表。
export function formulaPrompt({ intent, context = {}, lang }) {
  const headers = (Array.isArray(context.headers) ? context.headers : []).slice(0, 40).map(h => `${String(h.col || '').slice(0, 4)}=${String(h.name || '').slice(0, 40)}`).join('，');
  const rows = (Array.isArray(context.rows) ? context.rows : []).slice(0, 5).map(r => (Array.isArray(r) ? r : []).slice(0, 40).map(v => String(v ?? '').slice(0, 40)).join(' | ')).join('\n');
  const info = `工作表：${String(context.sheet || '').slice(0, 60)}\n数据区域：${String(context.range || '').slice(0, 40)}（第一行是表头）\n结果写在：${String(context.target || '').slice(0, 20)}\n各列：${headers || '（未提供）'}\n示例数据：\n${rows || '（未提供）'}`;
  const sys = '你是 Excel / WPS 表格公式专家。根据需求写一个公式，只输出 JSON：{"formula":"=...","explain":"一句话说明"}。'
    + '规则：只用 Excel 2016 以后和 WPS 都支持的常用函数；用给出的列地址引用数据，整列范围用绝对引用（如 $C$2:$C$100）；'
    + '不要引用其他工作簿；不要用 WEBSERVICE、FILTERXML、HYPERLINK、CALL、REGISTER 等联网或执行类函数；不确定时选最简单可靠的写法。'
    + (lang === 'ja' ? 'explain は日本語で書く。' : 'explain 用简体中文。');
  return [{ role: 'system', content: sys }, { role: 'user', content: `需求：${intent}\n\n${info}` }];
}
const BANNED = /\b(WEBSERVICE|FILTERXML|HYPERLINK|CALL|REGISTER(?:\.ID)?|EXEC|RTD|DDE|IMAGE)\s*\(/i;
export function parseFormula(answer) {
  const s = String(answer || '');
  let formula = '', explain = '';
  const a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a >= 0 && b > a) { try { const d = JSON.parse(s.slice(a, b + 1)); formula = String(d.formula || ''); explain = String(d.explain || ''); } catch {} }
  if (!formula) { const m = s.match(/=[^\n`]+/); if (m) formula = m[0]; }
  formula = formula.trim().replace(/^`+|`+$/g, '');
  if (!formula.startsWith('=')) throw Object.assign(Error('AI 没有给出公式，请把需求说得更具体些再试。'), { status: 422 });
  if (formula.length > 2000) throw Object.assign(Error('AI 给出的公式太长，请把需求拆小一点。'), { status: 422 });
  if (/[\r\n]/.test(formula)) formula = formula.replace(/\s*[\r\n]+\s*/g, '');
  if (BANNED.test(formula)) throw Object.assign(Error('AI 给出的公式用到了联网或执行类函数，为安全起见没有采用。'), { status: 422 });
  if (/\[[^\]]+\.(xlsx|xlsm|xls|et|csv)\]/i.test(formula) || /^=\s*[A-Za-z]+\|/.test(formula)) throw Object.assign(Error('AI 给出的公式引用了其他文件，为安全起见没有采用。'), { status: 422 });
  return { formula, explain: explain.slice(0, 300) };
}
