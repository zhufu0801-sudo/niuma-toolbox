param([switch]$NoOpen)
$ErrorActionPreference = 'Stop'
$appDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
# 报错提示的语言：设置里选的界面语言 > 日文版安装包 > Windows 显示语言
function Get-ToolboxLanguage {
  try {
    $saved = (Get-Content -LiteralPath (Join-Path $appDirectory 'generated\settings.json') -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json).ui.language
    if ($saved -eq 'ja' -or $saved -eq 'zh') { return $saved }
  } catch {}
  try { if ((Get-Content -LiteralPath (Join-Path $appDirectory 'lib\settings.mjs') -Raw -Encoding UTF8 -ErrorAction Stop) -match "language:'ja'") { return 'ja' } } catch {}
  try { if ((Get-UICulture).Name -like 'ja*') { return 'ja' } } catch {}
  return 'zh'
}
$isJa = (Get-ToolboxLanguage) -eq 'ja'
$texts = @{
  missing  = @('程序运行组件缺失，请重新安装办公工具箱。', '実行に必要なファイルが見つかりません。オフィスツールボックスをインストールし直してください。')
  port     = @('端口被其他程序占用。', 'ポートがほかのプログラムに使われています。')
  old      = @('本地服务版本较旧，请从已安装目录启动办公工具箱。', 'ローカルサービスのバージョンが古いです。インストール先のフォルダーから オフィスツールボックスを起動してください。')
  restart  = @('正在重新启动本应用的旧服务。', '古いサービスを再起動しています。')
  failed   = @('服务未能启动：', 'サービスを起動できませんでした：')
  slow     = @('服务启动比平时慢，请稍等几秒后再双击桌面上的“Office Toolbox”图标。', 'サービスの起動に時間がかかっています。数秒待ってから、デスクトップの「Office Toolbox」アイコンをもう一度ダブルクリックしてください。')
  wps      = @('启动页已打开，自动打开 WPS 未完成，请在页面点击“开始”。', 'スタートページを開きましたが、WPS を自動で開けませんでした。ページの「開始」をクリックしてください。')
  title    = @('办公工具箱', 'オフィスツールボックス')
}
function T([string]$key) { if ($isJa) { return $texts[$key][1] } return $texts[$key][0] }
try {
  $nodeRuntime = Join-Path $appDirectory 'runtime\node.exe'
  if (-not (Test-Path -LiteralPath $nodeRuntime)) { throw (T 'missing') }
  $serviceReady = $false
  try {
    $health = Invoke-RestMethod -Uri 'http://127.0.0.1:39127/api/health' -TimeoutSec 6
    if ($health.app -ne 'wps-ai-coach') { throw (T 'port') }
    $installedVersion = (Get-Content -LiteralPath (Join-Path $appDirectory 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json).version
    if ($health.version -ne $installedVersion) {
      $serverPath = Join-Path $appDirectory 'server.mjs'
      $ownedServices = Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -eq $nodeRuntime -and $_.CommandLine -and $_.CommandLine.Contains($serverPath) }
      if (-not $ownedServices) { throw (T 'old') }
      $ownedServices | ForEach-Object { Stop-Process -Id $_.ProcessId -ErrorAction Stop }
      throw (T 'restart')
    }
    $serviceReady = $true
  } catch {}
  if (-not $serviceReady) {
    $service = Start-Process -FilePath $nodeRuntime -ArgumentList ('"' + (Join-Path $appDirectory 'server.mjs') + '"') -WorkingDirectory $appDirectory -WindowStyle Hidden -RedirectStandardOutput (Join-Path $appDirectory 'service.log') -RedirectStandardError (Join-Path $appDirectory 'service-error.log') -PassThru
    # 重启电脑后第一次启动会慢很多（系统和杀毒软件会先检查运行组件），最多等 45 秒；服务进程提前退出就不再空等。
    for ($attempt = 0; $attempt -lt 90; $attempt++) {
      Start-Sleep -Milliseconds 500
      try { $health = Invoke-RestMethod -Uri 'http://127.0.0.1:39127/api/health' -TimeoutSec 3; if ($health.app -eq 'wps-ai-coach') { $serviceReady = $true; break } } catch {}
      if ($service -and $service.HasExited) { break }
    }
  }
  # 连续双击两次时，第二个服务会因端口已被占用而退出，但第一个其实已经在运行：报错前再确认一次。
  if (-not $serviceReady) { try { $health = Invoke-RestMethod -Uri 'http://127.0.0.1:39127/api/health' -TimeoutSec 3; if ($health.app -eq 'wps-ai-coach') { $serviceReady = $true } } catch {} }
  if (-not $serviceReady) {
    $detail = ''
    try { $detail = (Get-Content -LiteralPath (Join-Path $appDirectory 'service-error.log') -TotalCount 5 -ErrorAction Stop) -join "`n" } catch {}
    if ($detail) { throw ((T 'failed') + "`n" + $detail) }
    throw (T 'slow')
  }
  if (-not $NoOpen) {
    $startUrl = 'http://127.0.0.1:39127/start.html'
    $edgePaths = @(
      (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
      (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe'),
      (Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\Application\msedge.exe')
    )
    $edgeProgram = $edgePaths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if ($edgeProgram) { Start-Process -FilePath $edgeProgram -ArgumentList @('--app=' + $startUrl, '--window-size=800,700') -WindowStyle Normal }
    else { Start-Process -FilePath $startUrl }
    # 只有用户在设置中主动开启时，才同时打开 WPS。
    try {
      $session = Invoke-RestMethod -Uri 'http://127.0.0.1:39127/api/session' -TimeoutSec 5
      $headers = @{ 'X-Coach-Token' = $session.token }
      $preferences = Invoke-RestMethod -Uri 'http://127.0.0.1:39127/api/settings' -Headers $headers -TimeoutSec 5
      if ($preferences.settings.launch.autoOpenWps) { Invoke-RestMethod -Method Post -Uri 'http://127.0.0.1:39127/api/launch-wps' -Headers $headers -ContentType 'application/json' -Body '{}' -TimeoutSec 10 | Out-Null }
    } catch { Write-Warning (T 'wps') }
  }
} catch {
  if ($NoOpen) { throw }
  Add-Type -AssemblyName System.Windows.Forms
  [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, (T 'title')) | Out-Null
  exit 1
}
