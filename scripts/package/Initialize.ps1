param([switch]$RegisterOnly, [string]$Targets = '', [switch]$NoLaunch, [switch]$NoShortcut, [switch]$NoOfficeShortcuts, [string]$RegistrationPath = (Join-Path $env:APPDATA 'kingsoft\wps\jsaddons\publish.xml'))
# 办公工具箱安装：识别本机的 WPS Office 和 Microsoft Office，让用户选择装到哪个，再分别登记。
# -Targets 'wps','office' 或 'wps,office' 可跳过选择窗口（无人值守）；-NoShortcut 不建桌面快捷方式；-NoLaunch 装完不打开工具箱。
# 企业批量部署：在本文件旁边放 preset 文件夹（ai.json、settings.json、plugins\*.mjs），首次安装时复制进去；已有的文件不覆盖。
$ErrorActionPreference = 'Stop'
$appDirectory = Join-Path $PSScriptRoot 'app'
$generated = Join-Path $appDirectory 'generated'
$ja = (Get-Culture).Name -like 'ja*'
function Say($zh, $jp) { if ($ja) { $jp } else { $zh } }
function Find-Wps {
  foreach ($k in @('HKCU:\Software\Kingsoft\Office', 'HKLM:\SOFTWARE\WOW6432Node\Kingsoft\Office', 'HKLM:\SOFTWARE\Kingsoft\Office')) { if (Test-Path $k) { return $true } }
  return (Test-Path (Join-Path $env:LOCALAPPDATA 'Kingsoft\WPS Office')) -or (Test-Path (Join-Path ${env:ProgramFiles(x86)} 'Kingsoft\WPS Office'))
}
function Find-Office {
  $c = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
  if ($c -and $c.VersionToReport) {
    $ids = [string]$c.ProductReleaseIds
    $name = if ($ids -match 'O365|M365') { 'Microsoft 365' } elseif ($ids -match '2024') { 'Office 2024' } elseif ($ids -match '2021') { 'Office 2021' } elseif ($ids -match '2019') { 'Office 2019' } else { 'Microsoft Office' }
    # 快捷键需要共享运行时：Microsoft 365 / Office 2024 的 16.0.17928（2408）及以后。版本号形如 16.0.17928.20114。
    $build = 0; $parts = ([string]$c.VersionToReport).Split('.'); if ($parts.Length -ge 3) { [int]::TryParse($parts[2], [ref]$build) | Out-Null }
    return @{ found = $true; supported = ([string]$c.VersionToReport).StartsWith('16.'); name = $name; shortcuts = ($build -ge 17928) }
  }
  foreach ($b in @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office')) {
    foreach ($v in @('16.0', '15.0')) { foreach ($a in @('Excel', 'Word')) { $p = Get-ItemProperty ($b + '\' + $v + '\' + $a + '\InstallRoot') -ErrorAction SilentlyContinue; if ($p -and $p.Path) { return @{ found = $true; supported = ($v -eq '16.0'); name = $(if ($v -eq '16.0') { 'Office 2016' } else { 'Office 2013' }) } } } }
  }
  return @{ found = $false; supported = $false; name = '' }
}
function Choose-Targets($wps, $office) {
  Add-Type -AssemblyName System.Windows.Forms
  $form = New-Object System.Windows.Forms.Form
  $form.Text = Say '办公工具箱 · 选择要安装到的软件' 'オフィスツールボックス · インストール先の選択'
  $form.Width = 520; $form.Height = 270; $form.StartPosition = 'CenterScreen'; $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false
  $label = New-Object System.Windows.Forms.Label; $label.Left = 20; $label.Top = 15; $label.Width = 470; $label.Height = 40
  $label.Text = Say '已检查这台电脑。请选择要在哪个办公软件里使用工具箱（可以都选）：' 'このパソコンを確認しました。ツールボックスを使うソフトを選んでください（両方も可）：'
  $c1 = New-Object System.Windows.Forms.CheckBox; $c1.Left = 30; $c1.Top = 60; $c1.Width = 450
  $c1.Text = 'WPS Office' + $(if ($wps) { Say '（已检测到）' '（検出済み）' } else { Say '（未检测到）' '（見つかりません）' }); $c1.Checked = $wps
  $c2 = New-Object System.Windows.Forms.CheckBox; $c2.Left = 30; $c2.Top = 95; $c2.Width = 450
  $c2.Text = 'Microsoft Office' + $(if ($office.found) { '：' + $office.name + $(if ($office.supported) { Say '（已检测到，试用版）' '（検出済み・試用版）' } else { Say '（版本太旧，不支持）' '（古いため非対応）' }) } else { Say '（未检测到）' '（見つかりません）' })
  $c2.Checked = $office.found -and $office.supported; $c2.Enabled = $office.supported -or -not $office.found
  $note = New-Object System.Windows.Forms.Label; $note.Left = 30; $note.Top = 130; $note.Width = 460; $note.Height = 40
  $note.Text = Say 'Microsoft Office 版需要信任一个只在本机使用的安全证书，Windows 会弹出确认框，请选“是”。' 'Microsoft Office 版では、このパソコン専用のセキュリティ証明書を信頼します。確認画面が出たら「はい」を選んでください。'
  $ok = New-Object System.Windows.Forms.Button; $ok.Text = Say '安装' 'インストール'; $ok.Left = 300; $ok.Top = 180; $ok.Width = 90; $ok.DialogResult = 'OK'
  $cancel = New-Object System.Windows.Forms.Button; $cancel.Text = Say '取消' 'キャンセル'; $cancel.Left = 400; $cancel.Top = 180; $cancel.Width = 90; $cancel.DialogResult = 'Cancel'
  $form.Controls.AddRange(@($label, $c1, $c2, $note, $ok, $cancel)); $form.AcceptButton = $ok; $form.CancelButton = $cancel
  if ($form.ShowDialog() -ne 'OK') { throw (Say '已取消安装。' 'インストールを取り消しました。') }
  $list = @(); if ($c1.Checked) { $list += 'wps' }; if ($c2.Checked) { $list += 'office' }
  if (-not $list.Count) { throw (Say '没有选择任何软件，未安装。' 'ソフトが選ばれていないため、インストールしませんでした。') }
  return $list
}
function Register-Wps {
  $xml = New-Object System.Xml.XmlDocument
  $xml.PreserveWhitespace = $true
  if (Test-Path -LiteralPath $RegistrationPath) {
    $xml.Load($RegistrationPath)
    if ($xml.DocumentElement.Name -ne 'jsplugins') { throw 'Unexpected WPS registration format. Original file was not changed.' }
    Copy-Item -LiteralPath $RegistrationPath -Destination ($RegistrationPath + '.toolbox-backup-' + [Guid]::NewGuid().ToString('N'))
  } else {
    [void]$xml.AppendChild($xml.CreateXmlDeclaration('1.0', 'UTF-8', $null))
    [void]$xml.AppendChild($xml.CreateElement('jsplugins'))
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $RegistrationPath) | Out-Null
  }
  foreach ($entry in @(@('WpsCoachWriter','wps'), @('WpsCoachSheets','et'), @('WpsCoachSlides','wpp'))) {
    foreach ($old in @($xml.DocumentElement.SelectNodes("jspluginonline[@name='$($entry[0])']"))) { [void]$xml.DocumentElement.RemoveChild($old) }
    $item = $xml.CreateElement('jspluginonline')
    $item.SetAttribute('name', $entry[0])
    $item.SetAttribute('type', $entry[1])
    $item.SetAttribute('install', 'http://127.0.0.1:39127')
    $item.SetAttribute('url', ('http://127.0.0.1:39127/addon/' + $entry[1] + '/'))
    $item.SetAttribute('enable', 'enable_dev')
    [void]$xml.DocumentElement.AppendChild($item)
  }
  $temporary = $RegistrationPath + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
  $xml.Save($temporary)
  Move-Item -LiteralPath $temporary -Destination $RegistrationPath -Force
  Write-Host (Say '已登记到 WPS Office。' 'WPS Office に登録しました。')
}
function Register-Office {
  New-Item -ItemType Directory -Force -Path $generated | Out-Null
  $pfxPath = Join-Path $generated 'office-cert.pfx'; $keyPath = Join-Path $generated 'office-cert.key'; $thumbPath = Join-Path $generated 'office-cert.thumb'
  if (-not (Test-Path -LiteralPath $pfxPath)) {
    $cert = New-SelfSignedCertificate -DnsName 'localhost', '127.0.0.1' -CertStoreLocation 'Cert:\CurrentUser\My' -FriendlyName 'OfficeToolbox localhost' -NotAfter (Get-Date).AddYears(5) -KeyExportPolicy Exportable -KeyAlgorithm RSA -KeyLength 2048 -TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.1')
    $secret = [Guid]::NewGuid().ToString('N') + [Guid]::NewGuid().ToString('N')
    $secure = ConvertTo-SecureString $secret -AsPlainText -Force
    try { Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $secure -CryptoAlgorithmOption AES256_SHA256 | Out-Null } catch { Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $secure | Out-Null }
    Set-Content -LiteralPath $keyPath -Value $secret -Encoding ASCII
    Set-Content -LiteralPath $thumbPath -Value $cert.Thumbprint -Encoding ASCII
    $cer = Join-Path $generated 'office-cert.cer'; Export-Certificate -Cert $cert -FilePath $cer | Out-Null
    Write-Host (Say '正在信任本机证书：Windows 会弹出安全警告，请选“是”。' 'このパソコン専用の証明書を信頼します。セキュリティ警告が出たら「はい」を選んでください。')
    Import-Certificate -FilePath $cer -CertStoreLocation 'Cert:\CurrentUser\Root' | Out-Null
    Remove-Item -LiteralPath $cer
  }
  $idPath = Join-Path $generated 'office-id.txt'
  if (Test-Path -LiteralPath $idPath) { $id = (Get-Content -LiteralPath $idPath -Raw).Trim() } else { $id = [Guid]::NewGuid().ToString(); Set-Content -LiteralPath $idPath -Value $id -Encoding ASCII }
  $version = (Get-Content -LiteralPath (Join-Path $appDirectory 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json).version + '.0'
  $template = Get-Content -LiteralPath (Join-Path $appDirectory 'public\office\manifest-template.xml') -Raw -Encoding UTF8
  $manifest = $template.Replace('{{ID}}', $id).Replace('{{BASE}}', 'https://localhost:39128').Replace('{{VERSION}}', $version)
  $office = Find-Office
  if ($office.shortcuts -and -not $NoOfficeShortcuts) {
    $manifest = $manifest.Replace('{{REQUIREMENTS}}', '<Requirements><Sets DefaultMinVersion="1.1"><Set Name="SharedRuntime" MinVersion="1.1"/></Sets></Requirements>').Replace('{{RUNTIMES}}', '<Runtimes><Runtime resid="Taskpane.Url" lifetime="long"/></Runtimes>').Replace('{{EXTENDED}}', '<ExtendedOverrides Url="https://localhost:39128/office/shortcuts.json"></ExtendedOverrides>')
    Write-Host (Say '已开启 Office 快捷键：Ctrl+Alt+Shift+1 … 0（每个键执行哪个工具可在侧栏“快捷键”里改）。' 'Office のショートカットを有効にしました：Ctrl+Alt+Shift+1 … 0（各キーで実行するツールはサイドバーの「ショートカット」で変更できます）。')
  } else {
    $manifest = $manifest.Replace('{{REQUIREMENTS}}', '').Replace('{{RUNTIMES}}', '').Replace('{{EXTENDED}}', '')
  }
  $manifestPath = Join-Path $generated 'office-manifest.xml'
  [IO.File]::WriteAllText($manifestPath, $manifest, (New-Object System.Text.UTF8Encoding($false)))
  $key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
  New-Item -Path $key -Force | Out-Null
  Set-ItemProperty -Path $key -Name $id -Value $manifestPath
  Write-Host (Say '已登记到 Microsoft Office。请关闭并重新打开 Excel / Word，顶部会出现“办公工具箱”标签。' 'Microsoft Office に登録しました。Excel / Word を開き直すと「ツールボックス」タブが表示されます。')
}
function Copy-Preset {
  $preset = Join-Path $PSScriptRoot 'preset'
  if (-not (Test-Path -LiteralPath $preset)) { return }
  foreach ($name in @('ai.json', 'settings.json')) {
    $src = Join-Path $preset $name; $dst = Join-Path $generated $name
    if (-not (Test-Path -LiteralPath $src)) { continue }
    if (Test-Path -LiteralPath $dst) { Write-Host (Say ('已有 ' + $name + '，保留用户自己的设置，未覆盖。') ($name + ' がすでにあるため、利用者の設定を残しました（上書きしていません）。')); continue }
    try { Get-Content -LiteralPath $src -Raw -Encoding UTF8 | ConvertFrom-Json | Out-Null } catch { throw (Say ('preset\' + $name + ' 不是有效的 JSON，已停止安装。') ('preset\' + $name + ' は正しい JSON ではないため、インストールを止めました。')) }
    Copy-Item -LiteralPath $src -Destination $dst
    Write-Host (Say ('已应用企业预设：' + $name) ('企業向けの初期設定を適用しました：' + $name))
  }
  $plugins = Join-Path $preset 'plugins'
  if (Test-Path -LiteralPath $plugins) {
    $target = Join-Path $appDirectory 'plugins'; New-Item -ItemType Directory -Force -Path $target | Out-Null
    Get-ChildItem -LiteralPath $plugins -Filter '*.mjs' -File | ForEach-Object {
      $dst = Join-Path $target $_.Name
      if (Test-Path -LiteralPath $dst) { Write-Host (Say ('已有插件 ' + $_.Name + '，未覆盖。') ('プラグイン ' + $_.Name + ' はすでにあるため、上書きしていません。')) }
      else { Copy-Item -LiteralPath $_.FullName -Destination $dst; Write-Host (Say ('已放入隐私保护插件：' + $_.Name) ('個人情報保護プラグインを配置しました：' + $_.Name)) }
    }
  }
}
try {
  if (-not [Environment]::Is64BitOperatingSystem) { throw 'This package requires 64-bit Windows.' }
  $wps = Find-Wps; $office = Find-Office
  if ($Targets) { $chosen = @($Targets.Split(',') | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ }) }
  elseif ($RegisterOnly) { $chosen = @('wps') }
  else { $chosen = Choose-Targets $wps $office }
  New-Item -ItemType Directory -Force -Path $generated | Out-Null
  Copy-Preset
  if ($chosen -contains 'wps') { Register-Wps }
  if ($chosen -contains 'office') {
    Register-Office
    # 已在运行的旧服务没有加载证书：只停掉本目录启动的服务，下面的 Launch 会重新启动并开启 https。
    $nodeRuntime = Join-Path $appDirectory 'runtime\node.exe'; $serverPath = Join-Path $appDirectory 'server.mjs'
    Get-CimInstance Win32_Process | Where-Object { $_.ExecutablePath -eq $nodeRuntime -and $_.CommandLine -and $_.CommandLine.Contains($serverPath) } | ForEach-Object { Stop-Process -Id $_.ProcessId -ErrorAction SilentlyContinue }
  }
  [IO.File]::WriteAllText((Join-Path $generated 'targets.json'), (ConvertTo-Json @{ targets = $chosen; wpsFound = $wps; office = $office.name } -Compress), (New-Object System.Text.UTF8Encoding($false)))
  if ($RegisterOnly) { return }
  if (-not $NoShortcut) {
  $shell = New-Object -ComObject WScript.Shell
  $shortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'Office Toolbox.lnk'))
  $shortcut.TargetPath = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
  $shortcut.Arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + (Join-Path $appDirectory 'Launch.ps1') + '"'
  $shortcut.WorkingDirectory = $appDirectory
  $shortcut.IconLocation = Join-Path $appDirectory 'public\office-tools.ico'
  $shortcut.Save()
  }
  if ($NoLaunch) { Write-Host (Say '安装完成（未打开工具箱）。' 'インストールが完了しました（ツールボックスは開いていません）。'); return }
  & (Join-Path $appDirectory 'Launch.ps1')
  Write-Host (Say '完成。请保存文档，再重新打开 WPS 或 Excel / Word。' '完了しました。文書を保存して、WPS または Excel / Word を開き直してください。')
} catch { Write-Host $_.Exception.Message -ForegroundColor Red; exit 1 }
