// 安装包瘦身：删除运行时用不到的文件（源码映射、类型声明、说明文档、浏览器版构建）。
// 用法：node scripts/package/slim.mjs <安装包里的 app 文件夹>
// 只删下面列出的类别；许可证文件（LICENSE / NOTICE 等）一律保留。删完会打印节省的空间。
import fs from 'node:fs';
import path from 'node:path';

const FILE_RULES = [
  /\.map$/i,                       // 源码映射（调试用）
  /\.d\.(ts|mts|cts)$/i,           // TypeScript 类型声明
  /^(readme|changelog|history|contributing|code_of_conduct)(\.[a-z]+)?$/i,
  /\.(iife\.js|umd\.cjs|umd\.js)$/i // 浏览器 / 旧式打包版本
];
const KEEP = /^(licen[cs]e|notice|copying|authors)/i;
// 只在 node_modules 下按路径删除的整个文件夹（这些都不在 Node 运行路径上）
const DIRS = ['@types', 'pdfjs-dist/build', 'pdfjs-dist/web', 'pdfjs-dist/types', 'pdfjs-dist/image_decoders', 'pdf-parse/dist/pdf-parse/web', 'kuromoji/build', 'kuromoji/demo'];
const TEST_DIRS = new Set(['test', 'tests', '__tests__', 'example', 'examples', 'docs', '.github']);

export function slim(appDir, { dryRun = false } = {}) {
  const nm = path.join(appDir, 'node_modules');
  if (!fs.existsSync(nm)) throw Error('找不到 node_modules：' + nm);
  let bytes = 0, files = 0;
  const sizeOf = p => { const st = fs.lstatSync(p); if (!st.isDirectory()) return st.size; let s = 0; for (const n of fs.readdirSync(p)) s += sizeOf(path.join(p, n)); return s; };
  const remove = p => { bytes += sizeOf(p); files++; if (!dryRun) fs.rmSync(p, { recursive: true, force: true }); };
  for (const d of DIRS) { const p = path.join(nm, d); if (fs.existsSync(p)) remove(p); }
  (function walk(dir, depth) {
    for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, ent.name);
      if (ent.isDirectory()) { if (TEST_DIRS.has(ent.name.toLowerCase()) && depth >= 1 && fs.existsSync(path.join(dir, 'package.json'))) { remove(p); continue; } walk(p, depth + 1); continue; }
      if (KEEP.test(ent.name)) continue;
      if (FILE_RULES.some(r => r.test(ent.name))) remove(p);
    }
  })(nm, 0);
  return { bytes, files };
}

if (import.meta.url === `file://${process.argv[1]}` || process.argv[1]?.endsWith('slim.mjs')) {
  const dir = process.argv[2];
  if (!dir) { console.error('用法：node scripts/package/slim.mjs <app 文件夹> [--dry-run]'); process.exit(2); }
  const r = slim(path.resolve(dir), { dryRun: process.argv.includes('--dry-run') });
  console.log(`已删除 ${r.files} 项，节省 ${(r.bytes / 1048576).toFixed(1)} MB${process.argv.includes('--dry-run') ? '（演练，未实际删除）' : ''}`);
}
