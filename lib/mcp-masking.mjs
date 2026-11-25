// Per-MCP-process mappings. Never persist document text or mappings.
import {maskEngine} from './masking.mjs';
const TTL = 2 * 60 * 60 * 1000;
const strings = (value, fn) => typeof value === 'string' ? fn(value) : Array.isArray(value) ? value.map(v => strings(v, fn)) : value && typeof value === 'object' ? Object.fromEntries(Object.entries(value).map(([k,v]) => [k,strings(v,fn)])) : value;
export function createMcpMasker({vault, execute, getLanguage = () => 'zh', now = Date.now, ttl = TTL}) {
  const sessions = new Map();
  function sweep() { for (const [id,s] of sessions) if (now()-s.at >= ttl) sessions.delete(id); }
  const timer = setInterval(sweep, Math.min(ttl,60000)); timer.unref?.();
  return {
    close() { clearInterval(timer); sessions.clear(); },
    async execute(input) {
      sweep();
      const prefs = await vault.prefs();
      if (prefs.agent === false) return execute(input);
      // Reject old/unknown callers rather than mixing all callers into a shared mapping.
      const id = input.sessionId;
      if (typeof id !== 'string' || !/^[a-zA-Z0-9-]{16,80}$/.test(id)) throw Error('请重启 MCP 连接以启用会话脱敏。MCP 接続を再起動してください。');
      let s = sessions.get(id);
      if (!s) { if (sessions.size >= 256) throw Error('MCP 会话过多，请稍后重试。MCP セッションが多すぎます。'); s = {at:now(),pairs:[]}; sessions.set(id,s); }
      s.at = now();
      const E = await maskEngine(), unknown = new Set();
      const restore = text => { const r = E.restoreCodes(text,s.pairs); r.unknown.forEach(x=>unknown.add(x)); return r.text; };
      let request = input;
      if (input.action === 'insertText') request = {...input,args:{...input.args,text:typeof input.args?.text==='string'?restore(input.args.text):input.args?.text}};
      if (input.action === 'runTool') request = {...input,args:{...input.args,options:strings(input.args?.options,restore)}};
      const words = prefs.types.includes('WORD') ? E.parseWords(await vault.currentWords()).words.map(w=>({word:w.word,code:''})) : [];
      const lang = (await getLanguage()) === 'ja' ? 'ja' : 'zh';
      // Keep previously issued codes stable, even after the UI language changes.
      const session = E.session({mode:'code',lang,words,prior:s.pairs});
      let count = 0;
      const mask = text => { const r = E.maskText(text,session,{types:prefs.types,words}); count += r.seq.length; return r.text; };
      let result;
      try { result = await execute(request); }
      catch (e) { const message = mask(String(e.message||e)); s.pairs.push(...session.pairs()); throw Object.assign(Error(message),{status:e.status}); }
      // Include returned messages too: a host may echo the text just inserted.
      result = strings(result,mask);
      s.pairs.push(...session.pairs()); s.at=now();
      return {...(result && typeof result==='object' && !Array.isArray(result)?result:{result}),masking:{enabled:true,count,unknown:[...unknown],notice:`已把 ${count} 处个人信息换成代号，写回时请原样保留代号。個人情報 ${count} か所を記号に置き換えました。書き戻すときは記号をそのまま残してください。`,...(unknown.size?{warning:'不认识的代号原样保留。対応表にない記号はそのまま残しました。'}:{})}};
    }
  };
}
