$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$nodeCommand = Get-Command node.exe -ErrorAction SilentlyContinue
if (-not $nodeCommand) {
  $nodeCandidate = Join-Path $env:LOCALAPPDATA 'Programs\node-v24.22.2-win-x64\node.exe'
  if (Test-Path -LiteralPath $nodeCandidate) { $nodePath = $nodeCandidate } else { throw '请先安装 Node.js 22 或更新版本。' }
} else { $nodePath = $nodeCommand.Source }
$ollamaPath = Join-Path $env:LOCALAPPDATA 'Programs\Ollama\ollama.exe'
try { Invoke-RestMethod -Uri 'http://127.0.0.1:11434/api/version' -TimeoutSec 2 | Out-Null } catch {
  if (Test-Path -LiteralPath $ollamaPath) { Start-Process -FilePath $ollamaPath -ArgumentList 'serve' -WindowStyle Hidden }
}
$serviceReady = $false
try { $health = Invoke-RestMethod -Uri 'http://127.0.0.1:39127/api/health' -TimeoutSec 4; $serviceReady = $health.app -eq 'wps-ai-coach' } catch {}
if (-not $serviceReady) {
  $serverFile = Join-Path $projectDir 'server.mjs'
  Start-Process -FilePath $nodePath -ArgumentList ('"' + $serverFile + '"') -WorkingDirectory $projectDir -WindowStyle Hidden -RedirectStandardOutput (Join-Path $projectDir 'service.log') -RedirectStandardError (Join-Path $projectDir 'service-error.log')
  for ($attempt = 0; $attempt -lt 20; $attempt++) {
    Start-Sleep -Milliseconds 300
    try { $health = Invoke-RestMethod -Uri 'http://127.0.0.1:39127/api/health' -TimeoutSec 2; if ($health.app -eq 'wps-ai-coach') { $serviceReady = $true; break } } catch {}
  }
}
if (-not $serviceReady) { throw '本地服务未启动，请查看 service-error.log。' }
Start-Process 'http://127.0.0.1:39127/'
