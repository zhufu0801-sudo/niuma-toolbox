// 更新安装包里的日本邮编数据（lib/data/postal-jp.json.gz）。发布新版本前运行一次（日本邮政每月更新）。
//   node scripts/data/update-postal-jp.mjs                 从日本邮政官网下载 utf_ken_all.zip
//   node scripts/data/update-postal-jp.mjs --file 路径      用自己下载好的 utf_ken_all.zip（或解压后的 CSV）
//   数据日期取 zip 里文件的日期（日本邮政更新那天），不是运行这个脚本的日子。
import { downloadOfficial, dataFromFile, writeData, BUNDLED, createPostal } from '../../lib/postal-jp.mjs';
const i = process.argv.indexOf('--file'), file = i > 0 ? process.argv[i + 1] : '';
const data = file ? await dataFromFile(file) : await downloadOfficial();
const p = createPostal(data), check = p.byZip('100-0001');
if (!check.ok) throw Error('数据里查不到 100-0001（千代田区千代田），数据可能不完整，没有写入。');
await writeData(data, BUNDLED);
console.log(`已更新 ${BUNDLED}：${data.rows.length} 条，市区町村 ${data.cities.length} 个，数据日期 ${data.date}。`);
