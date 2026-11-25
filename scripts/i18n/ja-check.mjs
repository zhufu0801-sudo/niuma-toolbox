// 日文翻译检查（在项目根目录运行）。
//   node scripts/i18n/ja-check.mjs extract public/tools/batch.js public/file-tools-panel.js …
//     列出这些文件里还没有日文翻译的中文字符串（拼接用的片段也会列出，判断时以完整消息为准）。
//   node scripts/i18n/ja-check.mjs samples 样例.json
//     样例.json 是完整消息的数组（运行时真正显示的文字，可含换行），逐条翻译，列出没翻译或残留简体字的。
import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
const root=process.cwd(),ctx={};ctx.window=ctx;ctx.document={documentElement:{},addEventListener(){},querySelectorAll:()=>[]};ctx.localStorage={getItem:()=>'ja',setItem(){}};ctx.navigator={language:'ja'};ctx.location={search:'?lang=ja'};ctx.setTimeout=()=>0;ctx.MutationObserver=function(){this.observe=()=>{};};
for(const f of ['i18n-ja.js','i18n-ja-extra.js','i18n.js'])vm.runInNewContext(fs.readFileSync(path.join(root,'public',f),'utf8'),ctx);
const I=ctx.OfficeI18n;try{I.set('ja');}catch{}
const [mode,...args]=process.argv.slice(2);
// 只认简体中文才有、日文汉字里没有的字，避免把日文汉字误判为残留。
const SIMP=/[这们页处张没请删设图标题选项对齐据为线码击键开关档转换样顶间个还进么发现给从无让务后实单报价货择简节复制编动结类吗呢该钮框栏读写认识错误汇总扩辑仅侧览确预]/;
if(mode==='extract'){
  // .js 取引号里的字符串；.html 另外取标签之间的文字和 placeholder／title／alt 属性（2026-11-20：工作室页面的半翻译就是 HTML 正文里漏的）。
  // 判定“没翻好”：完全没变，或者翻完还残留简体字（逐词拼出来的中日混杂）。
  const CJK=/[一-鿿]/,all=new Set(),unhtml=t=>t.replace(/&nbsp;/g,' ').replace(/&#10;/g,'\n').replace(/&#9;/g,'\t').replace(/&gt;/g,'>').replace(/&lt;/g,'<').replace(/&quot;/g,'"').replace(/&amp;/g,'&');
  for(const f of args){let s=fs.readFileSync(path.join(root,f),'utf8');
    if(/\.html?$/i.test(f)){const body=s.replace(/<script[\s\S]*?<\/script>/g,'').replace(/<style[\s\S]*?<\/style>/g,'');for(const m of body.matchAll(/>([^<>]+)</g)){const t=unhtml(m[1]).trim();if(CJK.test(t))all.add(t);}for(const m of body.matchAll(/(?:placeholder|title|aria-label|alt)="([^"]+)"/g)){const t=unhtml(m[1]).trim();if(CJK.test(t))all.add(t);}}
    if(!/\.html?$/i.test(f))for(const m of s.matchAll(/'((?:[^'\\\n]|\\.)*)'|"((?:[^"\\\n]|\\.)*)"/g)){const t=(m[1]??m[2]).replace(/\\n/g,'\n').replace(/\\'/g,"'");if(CJK.test(t)&&!/[<>]/.test(t))all.add(t);}}
  const miss=[...all].filter(t=>{const x=I.t(t);return (x===t&&!/[\u3040-\u30ff]/.test(t)&&!/^[＋+\s]*文字$/.test(t))||SIMP.test(x);});console.log(miss.length+' 条没有翻译或翻译不完整（共 '+all.size+' 条中文字符串）');for(const t of miss)console.log(JSON.stringify(t),'→',JSON.stringify(I.t(t)));
}else if(mode==='samples'){
  let bad=0;const list=JSON.parse(fs.readFileSync(args[0],'utf8'));
  for(const s of list){const t=I.t(s);if(t===s||SIMP.test(t)){bad++;console.log('✗',JSON.stringify(s),'\n   →',JSON.stringify(t));}}
  console.log(list.length+' 条样例，'+bad+' 条有问题（人名、文件路径里的中文属正常）');process.exitCode=bad?1:0;
}else console.log('用法：node scripts/i18n/ja-check.mjs extract <文件…> | samples <样例.json>');
