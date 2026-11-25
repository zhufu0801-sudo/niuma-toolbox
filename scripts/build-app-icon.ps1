$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$bitmap = New-Object System.Drawing.Bitmap(64,64)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$graphics.Clear([System.Drawing.Color]::Transparent)
$background = New-Object System.Drawing.Drawing2D.GraphicsPath
$background.AddArc(2,2,24,24,180,90)
$background.AddArc(38,2,24,24,270,90)
$background.AddArc(38,38,24,24,0,90)
$background.AddArc(2,38,24,24,90,90)
$background.CloseFigure()
$brush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#087F8C'))
$graphics.FillPath($brush,$background)
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::White,3)
$pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
$pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$graphics.DrawRectangle($pen,16,25,32,24)
$graphics.DrawLine($pen,25,25,25,18)
$graphics.DrawLine($pen,25,18,39,18)
$graphics.DrawLine($pen,39,18,39,25)
$graphics.DrawLine($pen,16,34,48,34)
$graphics.DrawRectangle($pen,29,32,6,7)
$icon = [System.Drawing.Icon]::FromHandle($bitmap.GetHicon())
$outputPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'public\office-tools.ico'
$stream = [System.IO.File]::Create($outputPath)
try { $icon.Save($stream) } finally { $stream.Dispose(); $icon.Dispose(); $graphics.Dispose(); $bitmap.Dispose(); $pen.Dispose(); $brush.Dispose(); $background.Dispose() }
Write-Output "Generated desktop icon: $outputPath"
