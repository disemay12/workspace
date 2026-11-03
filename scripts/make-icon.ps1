# Генерация иконки приложения: public/icon.png (1024) и build/icon.ico (16…256).
# Рисуется кодом, чтобы иконка пересобиралась без внешних редакторов и зависимостей.

Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $PSScriptRoot
$publicDir = Join-Path $root 'public'
$buildDir = Join-Path $root 'build'
New-Item -ItemType Directory -Force -Path $publicDir, $buildDir | Out-Null

$size = 1024
$bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.Clear([System.Drawing.Color]::Transparent)

function New-RoundedPath([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
  $p = New-Object System.Drawing.Drawing2D.GraphicsPath
  $d = $r * 2
  $p.AddArc($x, $y, $d, $d, 180, 90)
  $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
  $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
  $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
  $p.CloseFigure()
  return $p
}

function New-SparkPath([float]$cx, [float]$cy, [float]$r) {
  $n = New-Object System.Drawing.PointF($cx, ($cy - $r))
  $e = New-Object System.Drawing.PointF(($cx + $r), $cy)
  $s = New-Object System.Drawing.PointF($cx, ($cy + $r))
  $w = New-Object System.Drawing.PointF(($cx - $r), $cy)
  $k = 0.12 * $r
  $c1 = New-Object System.Drawing.PointF(($cx + $k), ($cy - $k))
  $c2 = New-Object System.Drawing.PointF(($cx + $k), ($cy + $k))
  $c3 = New-Object System.Drawing.PointF(($cx - $k), ($cy + $k))
  $c4 = New-Object System.Drawing.PointF(($cx - $k), ($cy - $k))
  $p = New-Object System.Drawing.Drawing2D.GraphicsPath
  $p.AddBezier($n, $c1, $c1, $e)
  $p.AddBezier($e, $c2, $c2, $s)
  $p.AddBezier($s, $c3, $c3, $w)
  $p.AddBezier($w, $c4, $c4, $n)
  $p.CloseFigure()
  return $p
}

# --- Фон: скруглённый квадрат с градиентом ---
$margin = 26
$bgPath = New-RoundedPath $margin $margin ($size - $margin * 2) ($size - $margin * 2) 228
$bgBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
  (New-Object System.Drawing.Point(0, 0)),
  (New-Object System.Drawing.Point($size, $size)),
  ([System.Drawing.Color]::FromArgb(255, 30, 41, 59)),
  ([System.Drawing.Color]::FromArgb(255, 8, 12, 20)))
$g.FillPath($bgBrush, $bgPath)

# Мягкое свечение в правом верхнем углу: много тонких слоёв с шагом 2/255 —
# градиент без видимых колец и артефактов PathGradientBrush.
$glowState = $g.Save()
$g.SetClip($bgPath)
for ($i = 0; $i -lt 80; $i++) {
  $r = 430 - $i * 5
  $glowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(2, 16, 185, 129))
  $g.FillEllipse($glowBrush, (800 - $r), (245 - $r), ($r * 2), ($r * 2))
  $glowBrush.Dispose()
}
$g.Restore($glowState)

# Тонкая внутренняя рамка
$borderPath = New-RoundedPath ($margin + 3) ($margin + 3) ($size - $margin * 2 - 6) ($size - $margin * 2 - 6) 224
$borderPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(38, 255, 255, 255), 4)
$g.DrawPath($borderPen, $borderPath)

# --- Стопка документов ---
$cx = 500.0
$cy = 540.0
$docW = 360.0
$docH = 460.0

function Draw-Document([float]$angle, [string]$fill, [string]$border, [float]$shadowAlpha) {
  $state = $g.Save()
  $g.TranslateTransform($cx, $cy)
  $g.RotateTransform($angle)

  if ($shadowAlpha -gt 0) {
    $shadowPath = New-RoundedPath (-$docW / 2 + 10) (-$docH / 2 + 18) $docW $docH 34
    $shadowBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb([int]$shadowAlpha, 0, 0, 0))
    $g.FillPath($shadowBrush, $shadowPath)
    $shadowBrush.Dispose(); $shadowPath.Dispose()
  }

  $path = New-RoundedPath (-$docW / 2) (-$docH / 2) $docW $docH 34
  $brush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml($fill))
  $g.FillPath($brush, $path)
  if ($border) {
    $pen = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml($border), 3)
    $g.DrawPath($pen, $path)
    $pen.Dispose()
  }
  $brush.Dispose(); $path.Dispose()
  $g.Restore($state)
}

Draw-Document -22 '#243044' '#3b4a63' 90
Draw-Document -11 '#3d4c64' '#55647d' 80
Draw-Document 0 '#f7f9fc' '#ffffff' 70

# --- Содержимое переднего документа (координаты в его локальной системе) ---
$state = $g.Save()
$g.TranslateTransform($cx, $cy)

# Шапка документа
$header = New-RoundedPath (-134) (-168) 268 62 20
$headerBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#101826'))
$g.FillPath($headerBrush, $header)
$headerBrush.Dispose(); $header.Dispose()

# Три «окна» в шапке
$dots = @(
  @{ x = -100; y = -137; c = '#34d399' },
  @{ x = -68; y = -137; c = '#fbbf24' },
  @{ x = -36; y = -137; c = '#64748b' }
)
foreach ($dot in $dots) {
  $db = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml($dot.c))
  $g.FillEllipse($db, $dot.x, $dot.y, 22, 22)
  $db.Dispose()
}

# Строки текста
$lines = @(
  @{ y = -70; w = 232; c = '#0f172a' },
  @{ y = -18; w = 176; c = '#94a3b8' },
  @{ y = 34; w = 208; c = '#94a3b8' },
  @{ y = 86; w = 128; c = '#10b981' }
)
foreach ($line in $lines) {
  $lp = New-RoundedPath (-134) $line.y $line.w 22 11
  $lb = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml($line.c))
  $g.FillPath($lb, $lp)
  $lb.Dispose(); $lp.Dispose()
}

# Блок-карточка внизу документа
$card = New-RoundedPath (-134) 138 268 62 18
$cardBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#e2e8f0'))
$g.FillPath($cardBrush, $card)
$cardBrush.Dispose(); $card.Dispose()
$cardLine = New-RoundedPath (-112) 162 150 14 7
$cardLineBrush = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#94a3b8'))
$g.FillPath($cardLineBrush, $cardLine)
$cardLineBrush.Dispose(); $cardLine.Dispose()
$g.Restore($state)

# --- Искра AI ---
$sparkCx = 770.0
$sparkCy = 252.0
$sparkR = 104.0

$spark = New-SparkPath $sparkCx $sparkCy $sparkR
$sparkBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
  (New-Object System.Drawing.Point($sparkCx, ($sparkCy - $sparkR))),
  (New-Object System.Drawing.Point($sparkCx, ($sparkCy + $sparkR))),
  ([System.Drawing.Color]::FromArgb(255, 110, 231, 183)),
  ([System.Drawing.Color]::FromArgb(255, 16, 185, 129)))
$g.FillPath($sparkBrush, $spark)

# Маленькая искра-спутник
$mini = New-SparkPath 878 150 34
$miniBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(235, 167, 243, 208))
$g.FillPath($miniBrush, $mini)

$g.Dispose()

$pngPath = Join-Path $publicDir 'icon.png'
$bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

# --- ICO: несколько PNG-кадров (Vista+) ---
function Get-PngBytes([System.Drawing.Bitmap]$source, [int]$target) {
  $scaled = New-Object System.Drawing.Bitmap($target, $target, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $sg = [System.Drawing.Graphics]::FromImage($scaled)
  $sg.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
  $sg.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $sg.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $sg.Clear([System.Drawing.Color]::Transparent)
  $sg.DrawImage($source, 0, 0, $target, $target)
  $sg.Dispose()
  $ms = New-Object System.IO.MemoryStream
  $scaled.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
  $scaled.Dispose()
  return ,$ms.ToArray()
}

$iconSizes = @(16, 24, 32, 48, 64, 128, 256)
$frames = @()
foreach ($s in $iconSizes) {
  $frames += ,@($s, (Get-PngBytes $bmp $s))
}

$icoPath = Join-Path $buildDir 'icon.ico'
$stream = [System.IO.File]::Create($icoPath)
$writer = New-Object System.IO.BinaryWriter($stream)
$writer.Write([UInt16]0)
$writer.Write([UInt16]1)
$writer.Write([UInt16]$frames.Count)

$offset = 6 + 16 * $frames.Count
foreach ($frame in $frames) {
  $s = [int]$frame[0]
  $data = [byte[]]$frame[1]
  $dim = if ($s -ge 256) { 0 } else { $s }
  $writer.Write([byte]$dim)
  $writer.Write([byte]$dim)
  $writer.Write([byte]0)
  $writer.Write([byte]0)
  $writer.Write([UInt16]1)
  $writer.Write([UInt16]32)
  $writer.Write([UInt32]$data.Length)
  $writer.Write([UInt32]$offset)
  $offset += $data.Length
}
foreach ($frame in $frames) {
  $writer.Write([byte[]]$frame[1])
}
$writer.Flush()
$writer.Close()
$stream.Close()
$bmp.Dispose()

Write-Output "PNG: $pngPath"
Write-Output "ICO: $icoPath ($([math]::Round((Get-Item $icoPath).Length / 1KB, 1)) KB)"
