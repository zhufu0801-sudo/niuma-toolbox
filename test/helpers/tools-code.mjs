// 工具模块按页面里的加载顺序拼接，供 vm 测试一次性运行。顺序必须与 public/index.html、tools.html 一致。
import fs from 'node:fs';
export const TOOL_MODULES=['tools/core.js','tools/word.js','doc-outline.js','tools/doc-plus.js','tools/word-plus.js','tools/batch.js','tools/sheet.js','sheet-logic.js','tools/sheet-plus.js','tools/sheet-check.js','tools/slides.js','tools/transfer.js','tools/links.js','locale-engine.js','tools/locale.js','tools/print.js','date-engine.js','tools/dates.js','tools/approval.js','proof-engine.js','id-engine.js','tools/proof.js','tools/sheet-data.js','tools/slides-export.js','tools/files-more.js'];
export function toolsCode(){return TOOL_MODULES.map(f=>fs.readFileSync(new URL('../../public/'+f,import.meta.url),'utf8')).join('\n');}
