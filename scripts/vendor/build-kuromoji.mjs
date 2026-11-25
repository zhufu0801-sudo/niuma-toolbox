// 重新打包日文分词组件 kuromoji（只有升级 kuromoji 时才需要运行）。
// 为什么放进 lib/vendor 而不是依赖：装工具箱时不用再运行 npm install，离线电脑、Mac、GitHub 自动测试都一样能用。
// 用法（需要能上网装临时工具）：
//   npm install --no-save kuromoji@0.1.2 esbuild@0.24.0
//   node scripts/vendor/build-kuromoji.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..'), K = path.join(ROOT, 'node_modules', 'kuromoji'), OUT = path.join(ROOT, 'lib', 'vendor');
await build({ entryPoints: [path.join(K, 'src', 'kuromoji.js')], bundle: true, platform: 'node', format: 'cjs', minify: true, legalComments: 'none', outfile: path.join(OUT, 'kuromoji.min.cjs') });
fs.mkdirSync(path.join(OUT, 'kuromoji-dict'), { recursive: true });
for (const f of fs.readdirSync(path.join(K, 'dict'))) if (f.endsWith('.dat.gz')) fs.copyFileSync(path.join(K, 'dict', f), path.join(OUT, 'kuromoji-dict', f));
console.log('已更新 lib/vendor/kuromoji.min.cjs 和 lib/vendor/kuromoji-dict/。许可文件 lib/vendor/kuromoji-LICENSE.md 请对照新版本手动核对。');
