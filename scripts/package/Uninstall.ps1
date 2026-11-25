param([string]$RegistrationPath = (Join-Path $env:APPDATA 'kingsoft\wps\jsaddons\publish.xml'))
# 卸载办公工具箱：停止本目录的后台服务；从 WPS 和 Microsoft Office 移除登记；删除本机证书和桌面快捷方式。
# 不删除本文件夹，也不删除 app\generated 里的卡片和设置（AI 的 API Key 会删除）；确认不需要后，再手动删除整个文件夹。
$ErrorActionPreference = 'Stop'
$appDirectory = Join-Path $PSScriptRoot 'app'
$generated = Join-Path $appDirectory 'generated'
$ja = (Get-Culture).Name -like 'ja*'
function Say($zh, $jp) { if ($ja) { $jp } else { $zh } }
$messages = @()
try {
  $nodeRuntime = Join-Path $appDirectory 'runtime\node.exe'
  $serverPath = Join-Path $appDirectory 'server.mjs'
  $owned = @(Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -eq $nodeRuntime -and $_.CommandLine -and $_.CommandLine.Contains($serverPath) })
  foreach ($p in $owned) { Stop-Process -Id $p.ProcessId -ErrorAction SilentlyContinue }
  $messages += (Say '已停止后台服务：' 'バックグラウンドを停止しました：') + $owned.Count
  # WPS：只移除工具箱的三项，其他插件保留；改之前先备份
  if (Test-Path -LiteralPath $RegistrationPath) {
    $xml = New-Object System.Xml.XmlDocument
    $xml.PreserveWhitespace = $true
    $xml.Load($RegistrationPath)
    if ($xml.DocumentElement.Name -ne 'jsplugins') { throw (Say 'WPS 登记文件格式不认识，没有修改。' 'WPS の登録ファイルの形式を認識できないため、変更していません。') }
    $removed = 0
    foreach ($name in @('WpsCoachWriter','WpsCoachSheets','WpsCoachSlides')) {
      foreach ($node in @($xml.DocumentElement.SelectNodes("jspluginonline[@name='$name']"))) { [void]$xml.DocumentElement.RemoveChild($node); $removed++ }
    }
    if ($removed -gt 0) {
      Copy-Item -LiteralPath $RegistrationPath -Destination ($RegistrationPath + '.toolbox-uninstall-backup-' + (Get-Date -Format 'yyyyMMddHHmmss'))
      $temporary = $RegistrationPath + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
      $xml.Save($temporary)
      Move-Item -LiteralPath $temporary -Destination $RegistrationPath -Force
    }
    $messages += (Say '已从 WPS 移除登记：' 'WPS から登録を削除しました：') + $removed
  }
  # Microsoft Office：删除侧栏登记、本机证书和清单
  $idPath = Join-Path $generated 'office-id.txt'
  if (Test-Path -LiteralPath $idPath) {
    $id = (Get-Content -LiteralPath $idPath -Raw).Trim()
    Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer' -Name $id -ErrorAction SilentlyContinue
    $messages += (Say '已从 Microsoft Office 移除登记。' 'Microsoft Office から登録を削除しました。')
  }
  $thumbPath = Join-Path $generated 'office-cert.thumb'
  if (Test-Path -LiteralPath $thumbPath) {
    $thumb = (Get-Content -LiteralPath $thumbPath -Raw).Trim()
    foreach ($store in @('Cert:\CurrentUser\Root', 'Cert:\CurrentUser\My')) { $c = Get-Item -LiteralPath (Join-Path $store $thumb) -ErrorAction SilentlyContinue; if ($c) { Remove-Item -LiteralPath $c.PSPath } }
    foreach ($f in @('office-cert.pfx','office-cert.key','office-cert.thumb','office-manifest.xml')) { Remove-Item -LiteralPath (Join-Path $generated $f) -ErrorAction SilentlyContinue }
    $messages += (Say '已删除本机证书（Windows 可能弹出确认框）。' 'このパソコン用の証明書を削除しました（確認画面が出ることがあります）。')
  }
  # AI 的 API Key 属于账户密码，卸载时删除；卡片和其他设置保留
  $aiPath = Join-Path $generated 'ai.json'
  if (Test-Path -LiteralPath $aiPath) { Remove-Item -LiteralPath $aiPath; $messages += (Say '已删除保存的 AI 设置和 API Key。' '保存されていた AI 設定と API Key を削除しました。') }
  $shell = New-Object -ComObject WScript.Shell
  $desktop = [Environment]::GetFolderPath('Desktop')
  foreach ($name in @('Office Toolbox.lnk','办公工具箱.lnk')) {
    $path = Join-Path $desktop $name
    if (Test-Path -LiteralPath $path) {
      $link = $shell.CreateShortcut($path)
      if (($link.Arguments + ' ' + $link.TargetPath + ' ' + $link.WorkingDirectory).Contains($appDirectory)) { Remove-Item -LiteralPath $path; $messages += (Say '已删除桌面快捷方式：' 'デスクトップのショートカットを削除しました：') + $name }
    }
  }
  $messages += (Say '卸载完成。请保存文档后重新打开 WPS 或 Excel / Word，工具箱标签就会消失。' 'アンインストールが完了しました。文書を保存して WPS または Excel / Word を開き直すと、タブが消えます。')
  $messages += (Say '卡片和设置仍在：' 'カードと設定はここに残っています：') + $generated
  $messages | ForEach-Object { Write-Host $_ }
} catch { $messages | ForEach-Object { Write-Host $_ }; Write-Host $_.Exception.Message -ForegroundColor Red; exit 1 }
