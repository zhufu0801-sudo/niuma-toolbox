// 示例隐私插件（文件名以 _ 开头，不会被加载；改名为 example-privacy.mjs 即可启用）。
// 1. 内容里有“机密 / 極秘 / CONFIDENTIAL”时不发送。
// 2. 把公司内部项目代号替换成占位符，收到回答后换回。
const CODE_NAMES = ['星河计划', 'Project Aurora'];

export default {
  name: '示例：机密拦截与项目代号隐藏',
  version: '1.0.0',
  apiVersion: 1,
  async beforeSend(ctx) {
    const all = ctx.messages.map(m => m.content).join('\n');
    if (/机密|極秘|CONFIDENTIAL/i.test(all)) return { block: true, reason: '内容标有“机密”，按公司规定不能发给外部 AI。' };
    ctx.state.map = [];
    const messages = ctx.messages.map(m => {
      let content = m.content;
      CODE_NAMES.forEach((name, i) => {
        const token = `[PROJECT_${i + 1}]`;
        if (content.includes(name)) { content = content.split(name).join(token); ctx.state.map.push([token, name]); }
      });
      return { ...m, content };
    });
    return { messages };
  },
  async afterReceive(ctx) {
    let answer = ctx.answer;
    for (const [token, name] of ctx.state.map || []) answer = answer.split(token).join(name);
    return { answer };
  }
};
