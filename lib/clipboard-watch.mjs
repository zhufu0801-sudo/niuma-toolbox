// 读系统剪贴板里的文字：一次性读取（“从剪贴板添加”），或者持续留意复制动作（粘贴栏的“自动记录”）。
// Windows：一个常驻的 PowerShell，每 0.7 秒看一次剪贴板的“变更编号”，变了才读文字，不会反复占用剪贴板；
//   服务退出时一起结束，另外它每 5 秒检查一次服务进程还在不在，不在就自己退出。
// Mac：每 1.5 秒用 pbpaste 读一次，内容变了才算新复制（未实测）。
import { spawn, execFile } from 'node:child_process';

const WIN_WATCH = String.raw`
$ErrorActionPreference='SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
$cb=Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetClipboardSequenceNumber();' -Name ClipSeq -Namespace OfficeToolbox -PassThru
$parent=[int]$env:OFFICE_TOOLBOX_PID
$last=$cb::GetClipboardSequenceNumber();$tick=0
[Console]::Out.WriteLine('ready');[Console]::Out.Flush()
while($true){
  Start-Sleep -Milliseconds 700
  $tick++
  if($tick % 7 -eq 0 -and -not (Get-Process -Id $parent -ErrorAction SilentlyContinue)){exit}
  $n=$cb::GetClipboardSequenceNumber()
  if($n -ne $last){
    $last=$n
    Start-Sleep -Milliseconds 120
    try{ if([System.Windows.Forms.Clipboard]::ContainsText()){ $t=[System.Windows.Forms.Clipboard]::GetText(); if($t.Length -le 20000){ [Console]::Out.WriteLine([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($t))); [Console]::Out.Flush() } } }catch{}
  }
}`;
const WIN_READ = String.raw`Add-Type -AssemblyName System.Windows.Forms;$t='';try{if([System.Windows.Forms.Clipboard]::ContainsText()){$t=[System.Windows.Forms.Clipboard]::GetText()}}catch{};[Console]::Out.Write([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($t)))`;

export function readClipboardText({ platform = process.platform, run = execFile } = {}) {
  return new Promise((resolve, reject) => {
    if (platform === 'win32') {
      run('powershell.exe', ['-NoProfile', '-STA', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', WIN_READ], { windowsHide: true, timeout: 8000, maxBuffer: 1 << 20 }, (e, out) => {
        if (e) return reject(Error('读取剪贴板失败，请稍后再试。'));
        resolve(Buffer.from(String(out).trim(), 'base64').toString('utf8'));
      });
    } else if (platform === 'darwin') {
      run('pbpaste', [], { timeout: 5000, maxBuffer: 1 << 20 }, (e, out) => e ? reject(Error('读取剪贴板失败，请稍后再试。')) : resolve(String(out)));
    } else reject(Error('这台电脑上不能读取剪贴板。'));
  });
}

// onText(text)：每次复制了新文字就调用一次。返回 {stop()}。
export function watchClipboard(onText, { platform = process.platform, spawnImpl = spawn, read = readClipboardText, interval = 1500, pid = process.pid } = {}) {
  if (platform === 'win32') {
    const child = spawnImpl('powershell.exe', ['-NoProfile', '-STA', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', WIN_WATCH], { windowsHide: true, stdio: ['ignore', 'pipe', 'ignore'], env: { ...process.env, OFFICE_TOOLBOX_PID: String(pid) } });
    let buffer = '', stopped = false;
    child.stdout.setEncoding('utf8');
    child.stdout.on('data', chunk => {
      buffer += chunk; let i;
      while ((i = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, i).trim(); buffer = buffer.slice(i + 1);
        if (!line || line === 'ready') continue;
        try { onText(Buffer.from(line, 'base64').toString('utf8')); } catch {}
      }
    });
    child.on('error', () => {});
    return { stop() { if (stopped) return; stopped = true; try { child.kill(); } catch {} }, child };
  }
  if (platform === 'darwin') {
    let last = null, busy = false;
    const timer = setInterval(async () => {
      if (busy) return; busy = true;
      try { const text = await read({ platform }); if (last !== null && text !== last && text.trim()) onText(text); last = text; } catch {} finally { busy = false; }
    }, interval);
    timer.unref?.();
    return { stop() { clearInterval(timer); } };
  }
  return null;
}
