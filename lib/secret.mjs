// API Key 加密保存：Windows 上用系统自带的 DPAPI（“当前用户”范围）加密，
// 只有同一台电脑上的同一个 Windows 用户能解开；文件被复制走也读不出密钥。
// macOS 存进“钥匙串”（只有当前 Mac 用户能读）；ai.json 里只留一个编号（protection='keychain'）。
// 其他系统（开发、测试环境）没有这些，按原样保存并在界面上说明。
import { execFile } from 'node:child_process';
import { randomUUID } from 'node:crypto';

const ENTROPY = 'OfficeToolbox.ai-key.v1';
// 输入通过环境变量传给 PowerShell，不出现在命令行里（任务管理器看不到）。
function powershell(script, input) {
  return new Promise((resolve, reject) => {
    execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script],
      { env: { ...process.env, TOOLBOX_SECRET_IN: input }, windowsHide: true, timeout: 20000, maxBuffer: 1 << 20 },
      (error, stdout) => error ? reject(error) : resolve(String(stdout).trim()));
  });
}
const PROTECT = `Add-Type -AssemblyName System.Security; $e=[Text.Encoding]::UTF8.GetBytes('${ENTROPY}'); $b=[Convert]::FromBase64String($env:TOOLBOX_SECRET_IN); [Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Protect($b,$e,'CurrentUser'))`;
const UNPROTECT = `Add-Type -AssemblyName System.Security; $e=[Text.Encoding]::UTF8.GetBytes('${ENTROPY}'); $b=[Convert]::FromBase64String($env:TOOLBOX_SECRET_IN); [Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Unprotect($b,$e,'CurrentUser'))`;

const KC_SERVICE = 'OfficeToolbox.ai-key';
function security(args) {
  return new Promise((resolve, reject) => {
    execFile('/usr/bin/security', args, { timeout: 20000, maxBuffer: 1 << 20 }, (error, stdout) => error ? reject(error) : resolve(String(stdout).replace(/\n$/, '')));
  });
}
// service / label：钥匙串里的名字（数据脱敏的对照表密钥用自己的名字，和 AI Key 分开）。
export function createSecretBox({ platform = process.platform, run = powershell, keychain = security, service = KC_SERVICE, label = '办公工具箱 AI Key' } = {}) {
  if (platform === 'darwin') return macBox(keychain, service, label);
  const cache = new Map();
  const available = platform === 'win32';
  return {
    available,
    // 返回 { value, protection }：protection 为 'dpapi' 或 'none'
    async protect(plain) {
      plain = String(plain || '');
      if (!plain) return { value: '', protection: 'none' };
      if (!available) return { value: plain, protection: 'none' };
      const out = await run(PROTECT, Buffer.from(plain, 'utf8').toString('base64'));
      if (!/^[A-Za-z0-9+/=]{20,}$/.test(out)) throw Error('系统加密没有返回有效结果。');
      cache.set(out, plain);
      return { value: out, protection: 'dpapi' };
    },
    async unprotect(value, protection) {
      if (!value) return '';
      if (protection !== 'dpapi') return String(value);
      if (cache.has(value)) return cache.get(value);
      if (!available) throw Error('这个 API Key 是在 Windows 上加密的，只能在原来的电脑和 Windows 用户下使用。请重新填写 Key。');
      let plain;
      try { plain = Buffer.from(await run(UNPROTECT, value), 'base64').toString('utf8'); }
      catch { throw Error('无法解开保存的 API Key（可能换了电脑或 Windows 用户）。请在设置里重新填写 Key。'); }
      cache.set(value, plain);
      return plain;
    }
  };
}

// macOS 钥匙串：add-generic-password -U 覆盖同一编号；读取用 find-generic-password -w。
function macBox(keychain, service = KC_SERVICE, label = '办公工具箱 AI Key') {
  const cache = new Map();
  return {
    available: true,
    async protect(plain) {
      plain = String(plain || '');
      if (!plain) return { value: '', protection: 'none' };
      // security -w prints non-ASCII passwords as hex. Store UTF-8 as ASCII
      // base64 and version the reference so legacy ASCII keys stay unchanged.
      const id = 'kc-v1-' + randomUUID();
      await keychain(['add-generic-password', '-U', '-a', id, '-s', service, '-l', label, '-w', Buffer.from(plain,'utf8').toString('base64')]);
      cache.set(id, plain);
      return { value: id, protection: 'keychain' };
    },
    async unprotect(value, protection) {
      if (!value) return '';
      if (protection === 'dpapi') throw Error('这个 API Key 是在 Windows 上加密的，Mac 上解不开。请在设置里重新填写 Key。');
      if (protection !== 'keychain') return String(value);
      if (cache.has(value)) return cache.get(value);
      let plain;
      try {
        plain = await keychain(['find-generic-password', '-a', String(value), '-s', service, '-w']);
        if(String(value).startsWith('kc-v1-')){
          const bytes=Buffer.from(plain,'base64');
          if(bytes.toString('base64')!==plain)throw Error('Invalid keychain encoding');
          plain=bytes.toString('utf8');
        }
      }
      catch { throw Error('钥匙串里找不到保存的 API Key。请在设置里重新填写 Key。'); }
      cache.set(value, plain);
      return plain;
    }
  };
}
