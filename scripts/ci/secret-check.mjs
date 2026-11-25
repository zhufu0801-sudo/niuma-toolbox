// GitHub 自动测试用：在真的 Windows / Mac 上检查系统加密能不能用（单元测试里用的是假的）。
//   Windows：DPAPI（PowerShell）；Mac：钥匙串（security 命令）。
// 检查 AI Key 和数据脱敏对照表：加密后文件里没有原文；换一个新的实例（没有缓存）也能解开。
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { createSecretBox } from '../../lib/secret.mjs';
import { createMaskVault } from '../../lib/masking.mjs';

const want = process.platform === 'win32' ? 'dpapi' : process.platform === 'darwin' ? 'keychain' : null;
if (!want) { console.log('这台机器没有系统加密（Linux），跳过。'); process.exit(0); }
const dir = await fs.mkdtemp(path.join(os.tmpdir(), '办公工具箱 加密检查-'));
try {
  const secret = 'sk-测试-' + Date.now();
  const box = createSecretBox(), sealed = await box.protect(secret);
  assert.equal(sealed.protection, want, '加密方式应该是 ' + want);
  assert.ok(!sealed.value.includes(secret), '加密结果里不应该有原文');
  assert.equal(await createSecretBox().unprotect(sealed.value, sealed.protection), secret, '新实例解不开');
  console.log('✓ AI Key：' + want);

  const file = path.join(dir, 'masking.json'), opts = { service: 'OfficeToolbox.masking-key.ci', label: '办公工具箱 CI 检查' };
  const vault = createMaskVault(file, { secrets: createSecretBox(opts) });
  await vault.saveProfiles({ profiles: [{ name: '客户', words: '张三\n北京某某科技有限公司' }] });
  const rec = await vault.addRecord({ kind: 'doc', host: 'wps', name: '合同-脱敏.docx', mode: 'code', counts: { WORD: 1 }, total: 1, pairs: [{ original: '张三', masked: '【名称A】', type: 'WORD', count: 1 }], data: { seq: { 'st1#0': [['【名称A】', '张三', 'WORD']] } } });
  const raw = await fs.readFile(file, 'utf8');
  assert.ok(!raw.includes('张三') && !raw.includes('北京某某'), '对照表文件里不应该有原文');
  const again = createMaskVault(file, { secrets: createSecretBox(opts) });
  assert.equal((await again.state()).protection, want);
  assert.equal((await again.getRecord(rec.id)).pairs[0].original, '张三', '新实例读不出对照表');
  assert.equal(await again.currentWords(), '张三\n北京某某科技有限公司');
  console.log('✓ 数据脱敏对照表：' + want + '，重新打开也能读');
} finally { await fs.rm(dir, { recursive: true, force: true }); }
