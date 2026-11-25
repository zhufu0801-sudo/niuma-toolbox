param(
  [Parameter(Mandatory = $true)][string]$PackageDir,
  [string]$Thumbprint = '',
  [string]$PfxPath = '',
  [string]$TimestampServer = '',
  [switch]$WhatIf
)
# 办公工具箱安装包签名（准备用，尚未购买证书，未实际签过）。
# 给安装包里的 PowerShell 脚本加 Authenticode 签名，并检查内置 node.exe 原有的签名。
# 用法：
#   演练（只列出要签的文件，不改动）：  .\Sign-Package.ps1 -PackageDir D:\OfficeToolbox-0.11.0-Windows-x64 -WhatIf
#   用已导入“当前用户\个人”证书库的证书：.\Sign-Package.ps1 -PackageDir <目录> -Thumbprint <指纹> -TimestampServer <CA 提供的时间戳地址>
#   用 .pfx 文件（会询问密码）：         .\Sign-Package.ps1 -PackageDir <目录> -PfxPath D:\cert.pfx -TimestampServer <地址>
# 注意：签名必须是打包的最后一步。签完再改脚本的任何一个字（包括换行），签名就会失效。
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath $PackageDir).Path
$files = @(Get-ChildItem -LiteralPath $root -Filter '*.ps1' -File) + @(Get-Item -LiteralPath (Join-Path $root 'app\Launch.ps1') -ErrorAction SilentlyContinue) | Where-Object { $_ }
if (-not $files.Count) { throw "在 $root 没有找到要签名的 .ps1 文件。" }
Write-Host "要签名的脚本（$($files.Count) 个）："; $files | ForEach-Object { Write-Host ('  ' + $_.FullName.Substring($root.Length + 1)) }
$node = Join-Path $root 'app\runtime\node.exe'
if (Test-Path -LiteralPath $node) {
  $ns = Get-AuthenticodeSignature -LiteralPath $node
  Write-Host ("node.exe 原有签名：" + $ns.Status + $(if ($ns.SignerCertificate) { '（' + $ns.SignerCertificate.Subject + '）' } else { '' }))
}
if ($WhatIf) { Write-Host '演练结束，没有改动任何文件。'; return }
if (-not $TimestampServer) { throw '请用 -TimestampServer 填写证书颁发机构提供的时间戳服务器地址（不加时间戳，证书过期后签名会失效）。' }
if ($Thumbprint) { $cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.Thumbprint -eq $Thumbprint.Replace(' ', '').ToUpper() } | Select-Object -First 1 }
elseif ($PfxPath) { $cert = Get-PfxCertificate -FilePath $PfxPath }
else { throw '请用 -Thumbprint 或 -PfxPath 指定代码签名证书。' }
if (-not $cert) { throw '没有找到可用于代码签名的证书，请确认指纹，或证书带有“代码签名”用途。' }
if (-not $cert.HasPrivateKey) { throw '证书没有私钥（硬件令牌请先插好并装好驱动），无法签名。' }
$bad = 0
foreach ($f in $files) {
  $r = Set-AuthenticodeSignature -LiteralPath $f.FullName -Certificate $cert -TimestampServer $TimestampServer -HashAlgorithm SHA256
  $check = Get-AuthenticodeSignature -LiteralPath $f.FullName
  Write-Host ('  ' + $f.Name + '：' + $check.Status)
  if ($check.Status -ne 'Valid') { $bad++ }
}
if ($bad) { throw "$bad 个文件签名后检查不通过，请不要发布这个包。" }
Write-Host '全部签名并检查通过。请重新压缩 ZIP，再在另一台电脑上解压检查一次。'
