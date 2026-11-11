# ============================================================
#  桌面便签 · Sticky Notes   (Windows / PowerShell 5.1)
#  · 置顶小纸条：随手写、翻页看、一键新建 / 删除，自动保存
#  · 拖顶部黄带去移动；Alt+← / Alt+→ 翻页，Alt+N 新建，Alt+Del 删除
# ============================================================
param([switch]$Trace)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:Dir      = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:DataPath = Join-Path $script:Dir 'notes.json'
$script:LogPath  = Join-Path $script:Dir 'sticky.log'
$script:TraceOn  = [bool]$Trace

function Write-Log {
    param([string]$m)
    try { Add-Content -LiteralPath $script:LogPath -Value ('[' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '] ' + $m) -Encoding UTF8 } catch { }
}
function Trace-Log { param([string]$m) if ($script:TraceOn) { Write-Log $m } }

trap { Write-Log ('TRAP ' + $_.Exception.Message); exit 1 }

# ---------------- 单实例 ----------------
$script:Mutex = $null
try {
    $script:Mutex = New-Object System.Threading.Mutex -ArgumentList $false, 'Local\StickyNotes_DSH_Single'
    $owned = $false
    try { $owned = $script:Mutex.WaitOne(0) }
    catch [System.Threading.AbandonedMutexException] { $owned = $true }
    if (-not $owned) { exit 0 }
} catch { }

# ---------------- 原生 API / 自绘面板 ----------------
if (-not ('WinNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WinNative {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr h);
    [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr h, IntPtr dc);
    [DllImport("gdi32.dll")] public static extern int GetDeviceCaps(IntPtr dc, int idx);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
}
'@
}

if (-not ('StickyPanel' -as [type])) {
    Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @'
using System;
using System.Drawing;
using System.Windows.Forms;
public class StickyPanel : Panel {
    public StickyPanel() {
        this.DoubleBuffered = true;
        this.SetStyle(ControlStyles.OptimizedDoubleBuffer |
                      ControlStyles.AllPaintingInWmPaint |
                      ControlStyles.UserPaint |
                      ControlStyles.ResizeRedraw, true);
        this.UpdateStyles();
        this.BackColor = Color.FromArgb(255, 253, 242);
    }
}
'@
}

# ---------------- 通用小工具 ----------------
function New-RoundPath {
    param([float]$X, [float]$Y, [float]$W, [float]$H, [float]$R)
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    if ($R -le 0.6) {
        $p.AddRectangle((New-Object System.Drawing.RectangleF -ArgumentList $X, $Y, $W, $H))
        return $p
    }
    $d = $R * 2
    $p.AddArc($X, $Y, $d, $d, 180, 90)
    $p.AddArc($X + $W - $d, $Y, $d, $d, 270, 90)
    $p.AddArc($X + $W - $d, $Y + $H - $d, $d, $d, 0, 90)
    $p.AddArc($X, $Y + $H - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function Fill-RoundRect {
    param([System.Drawing.Graphics]$G, [System.Drawing.Brush]$B,
          [float]$X, [float]$Y, [float]$W, [float]$H, [float]$R)
    $p = New-RoundPath $X $Y $W $H $R
    $G.FillPath($B, $p)
    $p.Dispose()
}

function New-Rect {
    param([double]$X, [double]$Y, [double]$W, [double]$H)
    New-Object System.Drawing.RectangleF -ArgumentList ([float]$X), ([float]$Y), ([float]$W), ([float]$H)
}

# ---------------- 版式（设备像素，随 DPI 缩放） ----------------
function Get-Layout {
    $k    = [double]$script:S.K
    $W    = [double]$script:S.WW
    $H    = [double]$script:S.HH
    $pad  = 18.0 * $k
    $band = 40.0 * $k
    $icon = 26.0 * $k
    $gap  = 8.0 * $k
    $iconY = ($band - $icon) / 2.0
    $xClose = $W - $pad - $icon
    $xDel   = $xClose - $icon - $gap
    $xNew   = $xDel - $icon - $gap
    $textX = $pad
    $textY = $band + 14.0 * $k
    $bottomH = 46.0 * $k
    $textW = $W - 2.0 * $pad
    $textH = $H - $textY - $bottomH + 2.0 * $k
    $pagerW = 110.0 * $k
    $prevX = $W - $pad - $pagerW
    $py = $H - $bottomH + ($bottomH - $icon) / 2.0
    @{
        W = $W; H = $H; K = $k
        Pad = $pad; Band = $band
        New   = (New-Rect $xNew   $iconY $icon $icon)
        Del   = (New-Rect $xDel   $iconY $icon $icon)
        Close = (New-Rect $xClose $iconY $icon $icon)
        Prev  = (New-Rect $prevX       $py $icon $icon)
        Page  = (New-Rect ($prevX + $icon) $py (58.0 * $k) $icon)
        Next  = (New-Rect ($prevX + $icon + 58.0 * $k) $py $icon $icon)
        Text  = (New-Rect $textX $textY $textW $textH)
    }
}

# ---------------- 缩放 / 字号 ----------------
function Apply-Layout {
    $S = $script:S
    if (-not $S.Form -or -not $S.Box) { return }
    $k = [double]$S.K
    $L = Get-Layout
    $S.Form.Size = New-Object System.Drawing.Size -ArgumentList ([int]$S.WW), ([int]$S.HH)
    $S.Box.Location = New-Object System.Drawing.Point -ArgumentList ([int]$L.Text.X), ([int]$L.Text.Y)
    $S.Box.Size = New-Object System.Drawing.Size -ArgumentList ([int]$L.Text.Width), ([int]$L.Text.Height)
    $S.Ph.Location = New-Object System.Drawing.Point -ArgumentList ([int]($L.Text.X + 2)), ([int]($L.Text.Y + 4))
    $S.Ph.Size = New-Object System.Drawing.Size -ArgumentList ([int]($L.Text.Width - 6)), ([int](40 * $k))
    $rp = New-RoundPath 0 0 ([float]$S.WW) ([float]$S.HH) ([float](14 * $k))
    $S.Form.Region = New-Object System.Drawing.Region -ArgumentList $rp
    $rp.Dispose()
    $S.Panel.Invalidate()
}

function Set-FontPx {
    param([double]$Px)
    $S = $script:S
    if ($null -eq $Px) { return }
    if ($Px -lt 9) { $Px = 9 }
    if ($Px -gt 30) { $Px = 30 }
    if ([math]::Abs($Px - [double]$S.FontPx) -lt 0.01) { return }
    $S.FontPx = $Px
    $k = [double]$S.K
    $nf = New-Object System.Drawing.Font -ArgumentList $script:FontFamily, ([float]($Px * $k)), ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)
    $S.FBox = $nf
    if ($S.Box) { $S.Box.Font = $nf }
    if ($S.Ph)  { $S.Ph.Font = $nf }
    Update-Scrollbars
    $S.Panel.Invalidate()
}

function Step-Font {
    param([int]$Dir)
    Set-FontPx ([double]$script:S.FontPx + $Dir)
}

function Get-ResizeZone {
    param([int]$X, [int]$Y)
    $S = $script:S
    $b = 7.0 * [double]$S.K
    $tol = 3.0 * [double]$S.K
    $W = [double]$S.WW
    $H = [double]$S.HH
    if ($X -lt -$tol -or $Y -lt -$tol -or $X -gt ($W + $tol) -or $Y -gt ($H + $tol)) { return '' }
    $left = $X -le $b
    $right = $X -ge ($W - $b)
    $top = $Y -le $b
    $bottom = $Y -ge ($H - $b)
    if (-not ($left -or $right -or $top -or $bottom)) { return '' }
    $z = ''
    if ($top) { $z = 'n' } elseif ($bottom) { $z = 's' }
    if ($left) { $z = $z + 'w' } elseif ($right) { $z = $z + 'e' }
    return $z
}

function Get-ResizeCursor {
    param([string]$Zone)
    switch ($Zone) {
        'n'  { return [System.Windows.Forms.Cursors]::SizeNS }
        's'  { return [System.Windows.Forms.Cursors]::SizeNS }
        'e'  { return [System.Windows.Forms.Cursors]::SizeWE }
        'w'  { return [System.Windows.Forms.Cursors]::SizeWE }
        'ne' { return [System.Windows.Forms.Cursors]::SizeNESW }
        'sw' { return [System.Windows.Forms.Cursors]::SizeNESW }
        'nw' { return [System.Windows.Forms.Cursors]::SizeNWSE }
        'se' { return [System.Windows.Forms.Cursors]::SizeNWSE }
        default { return [System.Windows.Forms.Cursors]::Default }
    }
}

function Resize-Apply {
    $S = $script:S
    $k = [double]$S.K
    $cp = [System.Windows.Forms.Cursor]::Position
    $dx = $cp.X - $script:resizeStart.X
    $dy = $cp.Y - $script:resizeStart.Y
    $r = $script:resizeRect
    $minW = [int][math]::Round(260 * $k)
    $minH = [int][math]::Round(170 * $k)
    $mode = [string]$script:resizeMode
    $x = [int]$r.X; $y = [int]$r.Y; $w = [int]$r.Width; $h = [int]$r.Height
    if ($mode.Contains('w')) {
        $w = [int]$r.Width - $dx
        if ($w -lt $minW) { $w = $minW }
        $x = [int]$r.X + [int]$r.Width - $w
    } elseif ($mode.Contains('e')) {
        $w = [int]$r.Width + $dx
        if ($w -lt $minW) { $w = $minW }
    }
    if ($mode.Contains('n')) {
        $h = [int]$r.Height - $dy
        if ($h -lt $minH) { $h = $minH }
        $y = [int]$r.Y + [int]$r.Height - $h
    } elseif ($mode.Contains('s')) {
        $h = [int]$r.Height + $dy
        if ($h -lt $minH) { $h = $minH }
    }
    if ($w -eq $S.WW -and $h -eq $S.HH -and $x -eq $S.Form.Left -and $y -eq $S.Form.Top) { return }
    $S.WW = $w
    $S.HH = $h
    $S.Form.SetBounds($x, $y, $w, $h)
    Apply-Layout
}

# ---------------- 数据 ----------------
function New-EmptyNote { [pscustomobject]@{ text = '' } }

function Load-Data {
    $d = @{ x = $null; y = $null; w = $null; h = $null; font = $null; topmost = $true; current = 0; notes = (New-Object System.Collections.ArrayList) }
    if (Test-Path -LiteralPath $script:DataPath) {
        try {
            $raw = Get-Content -LiteralPath $script:DataPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim().Length -gt 0) {
                $o = $raw | ConvertFrom-Json
                Trace-Log ('load raw current=' + $o.current + ' notes=' + (@($o.notes).Count))
                if ($o.PSObject.Properties['x']) { $d.x = [int]$o.x }
                if ($o.PSObject.Properties['y']) { $d.y = [int]$o.y }
                if ($o.PSObject.Properties['w']) { $d.w = [int]$o.w }
                if ($o.PSObject.Properties['h']) { $d.h = [int]$o.h }
                if ($o.PSObject.Properties['font']) { $d.font = [double]$o.font }
                if ($o.PSObject.Properties['topmost']) { $d.topmost = [bool]$o.topmost }
                if ($o.PSObject.Properties['current']) { $d.current = [int]$o.current }
                if ($o.PSObject.Properties['notes']) {
                    foreach ($n in @($o.notes)) {
                        [void]$d.notes.Add([pscustomobject]@{ text = [string]$n.text })
                    }
                }
            }
        } catch { Trace-Log ('load fail ' + $_.Exception.Message) }
    }
    if ($d.notes.Count -eq 0) { [void]$d.notes.Add((New-EmptyNote)) }
    if ($d.current -lt 0) { $d.current = 0 }
    if ($d.current -ge $d.notes.Count) { $d.current = $d.notes.Count - 1 }
    return $d
}

function To-Crlf {
    param([string]$s)
    if ($null -eq $s) { return '' }
    return (($s -replace "`r`n", "`n") -replace "`n", "`r`n")
}

function To-Lf {
    param([string]$s)
    if ($null -eq $s) { return '' }
    return ($s -replace "`r`n", "`n")
}

function Sync-Current {
    $S = $script:S
    if (-not $S.Form) { return }
    if ($S.Notes.Count -eq 0) { return }
    if ($S.Index -ge 0 -and $S.Index -lt $S.Notes.Count) {
        $S.Notes[$S.Index].text = (To-Lf ([string]$S.Box.Text))
    }
}

function Save-Data {
    try {
        $S = $script:S
        if (-not $S.Form) { return }
        Sync-Current
        $items = @()
        foreach ($n in $S.Notes) { $items += [pscustomobject]@{ text = [string]$n.text } }
        $o = [ordered]@{
            x       = [int]$S.Form.Left
            y       = [int]$S.Form.Top
            w       = [int][math]::Round($S.WW / $S.K)
            h       = [int][math]::Round($S.HH / $S.K)
            font    = [double]$S.FontPx
            topmost = [bool]$S.TopMost
            current = [int]$S.Index
            notes   = $items
        }
        ($o | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $script:DataPath -Encoding UTF8
        $S.Dirty = $false
        $S.LastSaved = (Get-Date -Format 'HH:mm')
    } catch { Trace-Log ('save fail ' + $_.Exception.Message) }
}

# ---------------- 交互动作 ----------------
function Update-Placeholder {
    $S = $script:S
    if (-not $S.Ph) { return }
    $S.Ph.Visible = (([string]$S.Box.Text).Length -eq 0)
}

function Update-Scrollbars {
    $S = $script:S
    if (-not $S.Box) { return }
    $txt = [string]$S.Box.Text
    $want = 'None'
    if ($txt.Length -gt 0) {
        $w = [int]($S.Box.ClientSize.Width - 8)
        if ($w -lt 40) { $w = 40 }
        $sz = [System.Windows.Forms.TextRenderer]::MeasureText($txt, $S.Box.Font,
                (New-Object System.Drawing.Size $w, 100000),
                [System.Windows.Forms.TextFormatFlags]::WordBreak)
        if ($sz.Height -gt ($S.Box.ClientSize.Height - 4)) { $want = 'Vertical' }
    }
    if ([string]$S.Box.ScrollBars -ne $want) {
        $ss = $S.Box.SelectionStart
        $S.Box.ScrollBars = $want
        $S.Box.SelectionStart = $ss
        $S.Box.SelectionLength = 0
        try { $S.Box.ScrollToCaret() } catch { }
    }
}

function Show-Current {
    $S = $script:S
    $S.Loading = $true
    try { $S.Box.Text = (To-Crlf ([string]$S.Notes[$S.Index].text)) } finally { $S.Loading = $false }
    Update-Placeholder
    Update-Scrollbars
    $S.Panel.Invalidate()
}
function Go-Index {
    param([int]$i)
    $S = $script:S
    Sync-Current
    if ($i -lt 0) { $i = 0 }
    if ($i -ge $S.Notes.Count) { $i = $S.Notes.Count - 1 }
    if ($i -eq $S.Index) { return }
    $S.Index = $i
    Show-Current
    Save-Data
}

function Add-Note {
    $S = $script:S
    Sync-Current
    $at = $S.Index + 1
    if ($at -lt 0) { $at = 0 }
    if ($at -gt $S.Notes.Count) { $at = $S.Notes.Count }
    $S.Notes.Insert($at, (New-EmptyNote))
    $S.Index = $at
    Show-Current
    Save-Data
    $S.Box.Focus()
}

function Remove-Note {
    $S = $script:S
    Sync-Current
    if ($S.Notes.Count -le 1) {
        if (([string]$S.Notes[0].text).Trim().Length -gt 0) {
            $r = [System.Windows.Forms.MessageBox]::Show($S.Form, '清空这张便签的全部内容？', '便签', 'YesNo', 'Question')
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        $S.Notes[0].text = ''
        $S.Index = 0
        Show-Current
        Save-Data
        return
    }
    $hasText = ([string]$S.Notes[$S.Index].text).Trim().Length -gt 0
    if ($hasText) {
        $msg = '确定删除第 ' + ($S.Index + 1) + ' 张便签？' + [Environment]::NewLine + '删除后无法恢复。'
        $r = [System.Windows.Forms.MessageBox]::Show($S.Form, $msg, '便签', 'YesNo', 'Question')
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    $S.Notes.RemoveAt($S.Index)
    if ($S.Index -ge $S.Notes.Count) { $S.Index = $S.Notes.Count - 1 }
    if ($S.Index -lt 0) { $S.Index = 0 }
    Show-Current
    Save-Data
}

# ---------------- 绘制 ----------------
function Draw-IconBg {
    param([System.Drawing.Graphics]$G, [System.Drawing.RectangleF]$R, [bool]$Hot, [bool]$Press)
    $k = [double]$script:S.K
    if ($Press) {
        $b = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(38, 0, 0, 0))
        Fill-RoundRect -G $G -B $b -X $R.X -Y $R.Y -W $R.Width -H $R.Height -R ([float](7 * $k))
        $b.Dispose()
    } elseif ($Hot) {
        $b = New-Object System.Drawing.SolidBrush -ArgumentList ([System.Drawing.Color]::FromArgb(24, 0, 0, 0))
        Fill-RoundRect -G $G -B $b -X $R.X -Y $R.Y -W $R.Width -H $R.Height -R ([float](7 * $k))
        $b.Dispose()
    }
}

function Draw-Plus {
    param([System.Drawing.Graphics]$G, [System.Drawing.RectangleF]$R, [System.Drawing.Color]$C)
    $k = [double]$script:S.K
    $pen = New-Object System.Drawing.Pen -ArgumentList $C, ([float](1.9 * $k))
    $cx = $R.X + $R.Width / 2.0
    $cy = $R.Y + $R.Height / 2.0
    $d = 5.5 * $k
    $G.DrawLine($pen, [float]($cx - $d), [float]$cy, [float]($cx + $d), [float]$cy)
    $G.DrawLine($pen, [float]$cx, [float]($cy - $d), [float]$cx, [float]($cy + $d))
    $pen.Dispose()
}

function Draw-Trash {
    param([System.Drawing.Graphics]$G, [System.Drawing.RectangleF]$R, [System.Drawing.Color]$C)
    $k = [double]$script:S.K
    $pen = New-Object System.Drawing.Pen -ArgumentList $C, ([float](1.7 * $k))
    $cx = $R.X + $R.Width / 2.0
    $cy = $R.Y + $R.Height / 2.0
    $G.DrawLine($pen, [float]($cx - 6.5 * $k), [float]($cy - 3.5 * $k), [float]($cx + 6.5 * $k), [float]($cy - 3.5 * $k))
    $G.DrawLine($pen, [float]($cx - 2.5 * $k), [float]($cy - 6.5 * $k), [float]($cx + 2.5 * $k), [float]($cy - 6.5 * $k))
    $G.DrawLine($pen, [float]($cx - 2.5 * $k), [float]($cy - 6.5 * $k), [float]($cx - 2.5 * $k), [float]($cy - 3.5 * $k))
    $G.DrawLine($pen, [float]($cx + 2.5 * $k), [float]($cy - 6.5 * $k), [float]($cx + 2.5 * $k), [float]($cy - 3.5 * $k))
    $G.DrawLine($pen, [float]($cx - 5.0 * $k), [float]($cy - 1.5 * $k), [float]($cx - 4.0 * $k), [float]($cy + 6.5 * $k))
    $G.DrawLine($pen, [float]($cx + 5.0 * $k), [float]($cy - 1.5 * $k), [float]($cx + 4.0 * $k), [float]($cy + 6.5 * $k))
    $G.DrawLine($pen, [float]($cx - 4.0 * $k), [float]($cy + 6.5 * $k), [float]($cx + 4.0 * $k), [float]($cy + 6.5 * $k))
    $pen.Dispose()
}

function Draw-Cross {
    param([System.Drawing.Graphics]$G, [System.Drawing.RectangleF]$R, [System.Drawing.Color]$C)
    $k = [double]$script:S.K
    $pen = New-Object System.Drawing.Pen -ArgumentList $C, ([float](1.7 * $k))
    $cx = $R.X + $R.Width / 2.0
    $cy = $R.Y + $R.Height / 2.0
    $d = 5.0 * $k
    $G.DrawLine($pen, [float]($cx - $d), [float]($cy - $d), [float]($cx + $d), [float]($cy + $d))
    $G.DrawLine($pen, [float]($cx - $d), [float]($cy + $d), [float]($cx + $d), [float]($cy - $d))
    $pen.Dispose()
}

function Draw-Chevron {
    param([System.Drawing.Graphics]$G, [System.Drawing.RectangleF]$R, [bool]$Left,
          [System.Drawing.Color]$C, [bool]$Enabled, [bool]$Hot)
    $k = [double]$script:S.K
    if (-not $Enabled) {
        $C = [System.Drawing.Color]::FromArgb(90, $C.R, $C.G, $C.B)
    }
    $pen = New-Object System.Drawing.Pen -ArgumentList $C, ([float](1.9 * $k))
    $cx = $R.X + $R.Width / 2.0
    $cy = $R.Y + $R.Height / 2.0
    $d = 5.0 * $k
    if ($Left) {
        $G.DrawLine($pen, [float]($cx + $d * 0.6), [float]($cy - $d), [float]($cx - $d * 0.6), [float]$cy)
        $G.DrawLine($pen, [float]($cx - $d * 0.6), [float]$cy, [float]($cx + $d * 0.6), [float]($cy + $d))
    } else {
        $G.DrawLine($pen, [float]($cx - $d * 0.6), [float]($cy - $d), [float]($cx + $d * 0.6), [float]$cy)
        $G.DrawLine($pen, [float]($cx + $d * 0.6), [float]$cy, [float]($cx - $d * 0.6), [float]($cy + $d))
    }
    $pen.Dispose()
}

function Draw-Sticky {
    param([System.Drawing.Graphics]$G)
    $S = $script:S
    if (-not $S.Palette) { return }
    $P  = $S.Palette
    $k  = [double]$S.K
    $L  = Get-Layout
    $W  = [double]$S.WW
    $H  = [double]$S.HH

    $G.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $G.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

    # 纸张
    $paper = New-Object System.Drawing.SolidBrush -ArgumentList $P.Paper
    $G.FillRectangle($paper, [float]0, [float]0, [float]$W, [float]$H)
    $paper.Dispose()

    # 顶部黄带
    $bandRect = New-Object System.Drawing.RectangleF -ArgumentList ([float]0), ([float]0), ([float]$W), ([float]$L.Band)
    $lg = New-Object System.Drawing.Drawing2D.LinearGradientBrush -ArgumentList $bandRect, $P.Band1, $P.Band2, ([System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
    $G.FillRectangle($lg, $bandRect)
    $lg.Dispose()

    $bandLine = New-Object System.Drawing.Pen -ArgumentList $P.BandLine, ([float](1.0 * $k))
    $G.DrawLine($bandLine, [float]0, [float]$L.Band, [float]$W, [float]$L.Band)
    $bandLine.Dispose()

    # 标题
    $titleX = [float]$L.Pad
    $titleRect = New-Object System.Drawing.RectangleF -ArgumentList $titleX, ([float](9 * $k)), ([float](120 * $k)), ([float](22 * $k))
    $titleBrush = New-Object System.Drawing.SolidBrush -ArgumentList $P.Ink
    $sfNear = New-Object System.Drawing.StringFormat
    $sfNear.Alignment = [System.Drawing.StringAlignment]::Near
    $sfNear.LineAlignment = [System.Drawing.StringAlignment]::Center
    $G.DrawString('便签', $S.FTitle, $titleBrush, $titleRect, $sfNear)
    $tw = $G.MeasureString('便签', $S.FTitle).Width
    $subRect = New-Object System.Drawing.RectangleF -ArgumentList ([float]($titleX + $tw + 10 * $k)), ([float](9 * $k)), ([float](160 * $k)), ([float](22 * $k))
    $dimBrush = New-Object System.Drawing.SolidBrush -ArgumentList $P.Dim
    $G.DrawString(('共 ' + $S.Notes.Count + ' 张'), $S.FSub, $dimBrush, $subRect, $sfNear)

    # 图标按钮
    $iconColor = $P.Icon
    $hot = [string]$S.Hover
    Draw-IconBg -G $G -R $L.New   -Hot ($hot -eq 'new')   -Press ($S.Pressed -eq 'new')
    Draw-Plus    -G $G -R $L.New   -C $iconColor
    Draw-IconBg -G $G -R $L.Del   -Hot ($hot -eq 'del')   -Press ($S.Pressed -eq 'del')
    Draw-Trash   -G $G -R $L.Del   -C $iconColor
    Draw-IconBg -G $G -R $L.Close -Hot ($hot -eq 'close') -Press ($S.Pressed -eq 'close')
    Draw-Cross   -G $G -R $L.Close -C $iconColor

    # 底栏
    $canvasY = [float]($H - 46.0 * $k)
    $sepPen = New-Object System.Drawing.Pen -ArgumentList ([System.Drawing.Color]::FromArgb(255, 236, 219, 168)), ([float](1.0 * $k))
    $G.DrawLine($sepPen, [float]$L.Pad, $canvasY, [float]($W - $L.Pad), $canvasY)
    $sepPen.Dispose()
    $footBrush = New-Object System.Drawing.SolidBrush -ArgumentList $P.Dim
    if ($S.Dirty) { $hint = '正在输入…' } else { $hint = ('自动保存 · ' + $S.LastSaved) }
    $hintRect = New-Object System.Drawing.RectangleF -ArgumentList ([float]$L.Pad), ([float]($canvasY + 14 * $k)), ([float](180 * $k)), ([float](20 * $k))
    $G.DrawString($hint, $S.FSub, $footBrush, $hintRect, $sfNear)
    $footBrush.Dispose()

    $canPrev = $S.Index -gt 0
    $canNext = $S.Index -lt ($S.Notes.Count - 1)
    Draw-Chevron -G $G -R $L.Prev -Left $true  -C $P.Ink -Enabled $canPrev -Hot ($hot -eq 'prev')
    Draw-Chevron -G $G -R $L.Next -Left $false -C $P.Ink -Enabled $canNext -Hot ($hot -eq 'next')

    $pageText = ('' + ($S.Index + 1) + ' / ' + $S.Notes.Count)
    $sfCenter = New-Object System.Drawing.StringFormat
    $sfCenter.Alignment = [System.Drawing.StringAlignment]::Center
    $sfCenter.LineAlignment = [System.Drawing.StringAlignment]::Center
    $pageBrush = New-Object System.Drawing.SolidBrush -ArgumentList $P.Ink
    $G.DrawString($pageText, $S.FPager, $pageBrush, $L.Page, $sfCenter)
    $pageBrush.Dispose()
    $sfCenter.Dispose()
    $sfNear.Dispose()

    # 边框
    $borderPen = New-Object System.Drawing.Pen -ArgumentList $P.Border, ([float](1.0 * $k))
    $bp = New-RoundPath ([float](0.5 * $k)) ([float](0.5 * $k)) ([float]($W - 1.0 * $k)) ([float]($H - 1.0 * $k)) ([float](14 * $k))
    $G.DrawPath($borderPen, $bp)
    $bp.Dispose()
    $borderPen.Dispose()
}

# ---------------- 命中测试 ----------------
function Hit-Test {
    param([int]$X, [int]$Y)
    $L = Get-Layout
    if ($L.New.Contains($X, $Y))   { return 'new' }
    if ($L.Del.Contains($X, $Y))   { return 'del' }
    if ($L.Close.Contains($X, $Y)) { return 'close' }
    if ($L.Prev.Contains($X, $Y) -and $script:S.Index -gt 0) { return 'prev' }
    if ($L.Next.Contains($X, $Y) -and $script:S.Index -lt ($script:S.Notes.Count - 1)) { return 'next' }
    return ''
}

function Invoke-Hit {
    param([string]$Key)
    switch ($Key) {
        'new'   { Add-Note }
        'del'   { Remove-Note }
        'close' { if ($script:S.Form) { $script:S.Form.Close() } }
        'prev'  { Go-Index ($script:S.Index - 1) }
        'next'  { Go-Index ($script:S.Index + 1) }
    }
}

# ---------------- 状态与字体 ----------------
$script:S = @{
    Form = $null; Panel = $null; Box = $null; Ph = $null
    K = 1.0; WW = 380; HH = 300
    Notes = @(); Index = 0
    TopMost = $true; Hover = ''; Pressed = ''
    Dirty = $false; Loading = $false
    FontPx = 14; ResizeMode = ''
    LastSaved = (Get-Date -Format 'HH:mm')
    Palette = $null
}

$script:S.Palette = @{
    Paper    = [System.Drawing.Color]::FromArgb(255, 253, 242)
    Band1    = [System.Drawing.Color]::FromArgb(251, 234, 160)
    Band2    = [System.Drawing.Color]::FromArgb(247, 226, 140)
    BandLine = [System.Drawing.Color]::FromArgb(232, 207, 138)
    Border   = [System.Drawing.Color]::FromArgb(227, 201, 107)
    Ink      = [System.Drawing.Color]::FromArgb(61, 54, 39)
    Dim      = [System.Drawing.Color]::FromArgb(160, 141, 78)
    Icon     = [System.Drawing.Color]::FromArgb(122, 100, 32)
}

# DPI
$k = 1.0
try {
    [WinNative]::SetProcessDPIAware() | Out-Null
    $hdc = [WinNative]::GetDC([IntPtr]::Zero)
    $dpi = [WinNative]::GetDeviceCaps($hdc, 88)
    [WinNative]::ReleaseDC([IntPtr]::Zero, $hdc) | Out-Null
    if ($dpi -ge 96) { $k = [double]$dpi / 96.0 }
} catch { $k = 1.0 }
if ($k -lt 1) { $k = 1.0 }
$script:S.K  = $k
$script:FontFamily = 'Microsoft YaHei UI'
Trace-Log ('dpi k=' + $k)

# ---------------- 数据 ----------------
$data = Load-Data
$script:S.Notes = $data.notes
$script:S.Index = $data.current
$script:S.TopMost = $data.topmost

# 尺寸 / 字号（用户可调，存于 notes.json；这里都是 96dpi 逻辑值）
$lw = 380
$lh = 300
$fp = 14
if ($null -ne $data.w)    { $lw = [int]$data.w }
if ($null -ne $data.h)    { $lh = [int]$data.h }
if ($null -ne $data.font) { $fp = [double]$data.font }
if ($lw -lt 260) { $lw = 260 }
if ($lh -lt 170) { $lh = 170 }
if ($fp -lt 9)   { $fp = 9 }
if ($fp -gt 30)  { $fp = 30 }
$wa0 = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
if (($lw * $k) -gt $wa0.Width)  { $lw = [int][math]::Floor($wa0.Width / $k) }
if (($lh * $k) -gt $wa0.Height) { $lh = [int][math]::Floor($wa0.Height / $k) }
$script:S.FontPx = $fp
$script:S.WW = [int][math]::Round($lw * $k)
$script:S.HH = [int][math]::Round($lh * $k)

$family = $script:FontFamily
$script:S.FTitle = New-Object System.Drawing.Font -ArgumentList $family, ([float](12 * $k)), ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
$script:S.FSub   = New-Object System.Drawing.Font -ArgumentList $family, ([float](10.5 * $k)), ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)
$script:S.FPager = New-Object System.Drawing.Font -ArgumentList $family, ([float](11.5 * $k)), ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
$script:S.FBox   = New-Object System.Drawing.Font -ArgumentList $family, ([float]($fp * $k)), ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)

# ---------------- 窗体 ----------------
$form = New-Object System.Windows.Forms.Form
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$form.StartPosition   = [System.Windows.Forms.FormStartPosition]::Manual
$form.ShowInTaskbar   = $false
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.ShowIcon        = $false
$form.Text            = '桌面便签'
$form.BackColor       = $script:S.Palette.Paper
$form.AutoScaleMode   = [System.Windows.Forms.AutoScaleMode]::None
$form.KeyPreview      = $true
$form.Size            = New-Object System.Drawing.Size -ArgumentList $script:S.WW, $script:S.HH
$form.TopMost         = $script:S.TopMost

$panel = New-Object StickyPanel
$panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($panel)

$regionPath = New-RoundPath 0 0 ([float]$script:S.WW) ([float]$script:S.HH) ([float](14 * $k))
$form.Region = New-Object System.Drawing.Region -ArgumentList $regionPath
$regionPath.Dispose()

# 文本框
$L0 = Get-Layout
$box = New-Object System.Windows.Forms.TextBox
$box.Multiline    = $true
$box.WordWrap     = $true
$box.ScrollBars   = [System.Windows.Forms.ScrollBars]::Vertical
$box.BorderStyle  = [System.Windows.Forms.BorderStyle]::None
$box.BackColor    = $script:S.Palette.Paper
$box.ForeColor    = $script:S.Palette.Ink
$box.Font         = $script:S.FBox
$box.Location     = New-Object System.Drawing.Point -ArgumentList ([int]$L0.Text.X), ([int]$L0.Text.Y)
$box.Size         = New-Object System.Drawing.Size -ArgumentList ([int]$L0.Text.Width), ([int]$L0.Text.Height)
$form.Controls.Add($box)

# 占位提示
$ph = New-Object System.Windows.Forms.Label
$ph.AutoSize   = $false
$ph.BackColor  = $script:S.Palette.Paper
$ph.ForeColor  = [System.Drawing.Color]::FromArgb(200, 184, 122)
$ph.Font       = $script:S.FBox
$ph.TextAlign  = [System.Drawing.ContentAlignment]::TopLeft
$ph.Text       = '在这里写下今天要记住的事…'
$ph.Location   = New-Object System.Drawing.Point -ArgumentList ([int]($L0.Text.X + 2)), ([int]($L0.Text.Y + 4))
$ph.Size       = New-Object System.Drawing.Size -ArgumentList ([int]($L0.Text.Width - 6)), ([int](30 * $k))
$form.Controls.Add($ph)
$box.BringToFront()
$ph.BringToFront()

$script:S.Form  = $form
$script:S.Panel = $panel
$script:S.Box   = $box
$script:S.Ph    = $ph
Apply-Layout

# 文本框内边距
try {
    $padPx = [int](10 * $k)
    $lp = [IntPtr](($padPx -shl 16) -bor $padPx)
    [void][WinNative]::SendMessage($box.Handle, 0x00D3, [IntPtr]3, $lp)
} catch { }

Show-Current

# ---------------- 位置 ----------------
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
if ($null -ne $data.x -and $null -ne $data.y) {
    $form.Left = [int]$data.x
    $form.Top  = [int]$data.y
} else {
    $form.Left = $wa.Right - $script:S.WW - [int](60 * $k)
    $form.Top  = $wa.Top + [int](($wa.Height - $script:S.HH) / 2)
}
$onScreen = $false
foreach ($sc in [System.Windows.Forms.Screen]::AllScreens) {
    $b = $sc.Bounds
    if ($form.Left -gt ($b.Left - 200) -and $form.Left -lt ($b.Right - 50) -and
        $form.Top  -gt ($b.Top  - 100) -and $form.Top  -lt ($b.Bottom - 40)) { $onScreen = $true; break }
}
if (-not $onScreen) {
    $form.Left = $wa.Right - $script:S.WW - [int](60 * $k)
    $form.Top  = $wa.Top + [int](($wa.Height - $script:S.HH) / 2)
}

# ---------------- 事件 ----------------
$script:dragActive = $false
$script:dragCursor = New-Object System.Drawing.Point 0, 0
$script:dragForm   = New-Object System.Drawing.Point 0, 0
$script:resizeMode = ''
$script:resizeStart = New-Object System.Drawing.Point 0, 0
$script:resizeRect  = New-Object System.Drawing.Rectangle 0, 0, 0, 0

$panel.Add_Paint({
    param($s, $e)
    try { Draw-Sticky -G $e.Graphics } catch { Write-Log ('paint: ' + $_.Exception.Message) }
})

$panel.Add_MouseMove({
    param($s, $e)
    if ($script:resizeMode) {
        if ([System.Windows.Forms.Control]::MouseButtons -ne [System.Windows.Forms.MouseButtons]::Left) {
            $script:resizeMode = ''
            Update-Scrollbars
            Save-Data
            return
        }
        Resize-Apply
        return
    }
    if ($script:dragActive) {
        if ([System.Windows.Forms.Control]::MouseButtons -ne [System.Windows.Forms.MouseButtons]::Left) {
            $script:dragActive = $false
            Save-Data
            return
        }
        $cp = [System.Windows.Forms.Cursor]::Position
        $script:S.Form.Left = $script:dragForm.X + ($cp.X - $script:dragCursor.X)
        $script:S.Form.Top  = $script:dragForm.Y + ($cp.Y - $script:dragCursor.Y)
        return
    }
    $zone = Get-ResizeZone -X $e.X -Y $e.Y
    if ($zone) {
        $script:S.Panel.Cursor = Get-ResizeCursor $zone
        if ($script:S.Hover -ne '') { $script:S.Hover = ''; $script:S.Panel.Invalidate() }
        return
    }
    $hit = Hit-Test -X $e.X -Y $e.Y
    if ($hit -ne $script:S.Hover) {
        $script:S.Hover = $hit
        if ($hit) { $script:S.Panel.Cursor = [System.Windows.Forms.Cursors]::Hand }
        else      { $script:S.Panel.Cursor = [System.Windows.Forms.Cursors]::Default }
        $script:S.Panel.Invalidate()
    }
})

$panel.Add_MouseLeave({
    $script:S.Pressed = ''
    if ($script:S.Hover -ne '') {
        $script:S.Hover = ''
        $script:S.Panel.Cursor = [System.Windows.Forms.Cursors]::Default
        $script:S.Panel.Invalidate()
    }
})

$panel.Add_MouseDown({
    param($s, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    $zone = Get-ResizeZone -X $e.X -Y $e.Y
    if ($zone) {
        $script:resizeMode = $zone
        $script:resizeStart = [System.Windows.Forms.Cursor]::Position
        $script:resizeRect = New-Object System.Drawing.Rectangle -ArgumentList $script:S.Form.Left, $script:S.Form.Top, $script:S.WW, $script:S.HH
        $script:S.Panel.Capture = $true
        return
    }
    $hit = Hit-Test -X $e.X -Y $e.Y
    if ($hit) {
        $script:S.Pressed = $hit
        $script:S.Panel.Invalidate()
        return
    }
    $script:dragActive = $true
    $script:dragCursor = [System.Windows.Forms.Cursor]::Position
    $script:dragForm = New-Object System.Drawing.Point -ArgumentList $script:S.Form.Left, $script:S.Form.Top
    $script:S.Panel.Capture = $true
})

$panel.Add_MouseUp({
    param($s, $e)
    $script:S.Panel.Capture = $false
    if ($script:resizeMode) {
        $script:resizeMode = ''
        Update-Scrollbars
        Save-Data
        return
    }
    if ($script:dragActive) {
        $script:dragActive = $false
        Save-Data
        return
    }
    $pressed = $script:S.Pressed
    if ($pressed) {
        $script:S.Pressed = ''
        $script:S.Panel.Invalidate()
        $hit = Hit-Test -X $e.X -Y $e.Y
        if ($hit -eq $pressed) { Invoke-Hit $pressed }
    }
})

$box.Add_TextChanged({
    $S = $script:S
    if ($S.Loading) { return }
    $S.Dirty = $true
    Update-Placeholder
    Update-Scrollbars
    $script:SaveTimer.Stop()
    $script:SaveTimer.Start()
    $S.Panel.Invalidate()
})

$box.Add_GotFocus({ Update-Placeholder; $script:S.Panel.Invalidate() })
$box.Add_LostFocus({ Update-Placeholder })

$ph.Add_Click({ $script:S.Box.Focus() })

$box.Add_MouseWheel({
    param($s, $e)
    if (([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control) -ne 0) {
        if ($e.Delta -gt 0) { Step-Font 1 } else { Step-Font -1 }
    }
})
$panel.Add_MouseWheel({
    param($s, $e)
    if (([System.Windows.Forms.Control]::ModifierKeys -band [System.Windows.Forms.Keys]::Control) -ne 0) {
        if ($e.Delta -gt 0) { Step-Font 1 } else { Step-Font -1 }
    }
})

$form.Add_KeyDown({
    param($s, $e)
    $kc = $e.KeyCode
    if ($e.Control) {
        if ($kc -eq [System.Windows.Forms.Keys]::Oemplus -or $kc -eq [System.Windows.Forms.Keys]::Add) {
            Step-Font 1
            $e.SuppressKeyPress = $true
        } elseif ($kc -eq [System.Windows.Forms.Keys]::OemMinus -or $kc -eq [System.Windows.Forms.Keys]::Subtract) {
            Step-Font -1
            $e.SuppressKeyPress = $true
        } elseif ($kc -eq [System.Windows.Forms.Keys]::D0) {
            Set-FontPx 14
            $e.SuppressKeyPress = $true
        }
        return
    }
    if (-not $e.Alt) { return }
    if ($kc -eq [System.Windows.Forms.Keys]::Left)  { Go-Index ($script:S.Index - 1); $e.SuppressKeyPress = $true }
    elseif ($kc -eq [System.Windows.Forms.Keys]::Right) { Go-Index ($script:S.Index + 1); $e.SuppressKeyPress = $true }
    elseif ($kc -eq [System.Windows.Forms.Keys]::N) { Add-Note; $e.SuppressKeyPress = $true }
    elseif ($kc -eq [System.Windows.Forms.Keys]::Delete) { Remove-Note; $e.SuppressKeyPress = $true }
})

# ---------------- 右键菜单 ----------------
$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miNew  = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '新建便签 (Alt+N)'
$miDel  = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '删除当前便签 (Alt+Del)'
$miPrev = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '上一张 (Alt+←)'
$miNext = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '下一张 (Alt+→)'
$miFontUp = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '字号 + (Ctrl+=)'
$miFontDn = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '字号 − (Ctrl+-)'
$miFontRe = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '字号重置 (Ctrl+0)'
$miSizeRe = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '大小重置 (380×300)'
$miCopy = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '复制 (Ctrl+C)'
$miPaste = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '粘贴 (Ctrl+V)'
$miAll  = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '全选 (Ctrl+A)'
$miTop  = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '始终置顶'
$miExit = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList '退出'
$miTop.Checked = $script:S.TopMost

[void]$menu.Items.Add($miNew)
[void]$menu.Items.Add($miDel)
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$menu.Items.Add($miPrev)
[void]$menu.Items.Add($miNext)
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$menu.Items.Add($miFontUp)
[void]$menu.Items.Add($miFontDn)
[void]$menu.Items.Add($miFontRe)
[void]$menu.Items.Add($miSizeRe)
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$menu.Items.Add($miCopy)
[void]$menu.Items.Add($miPaste)
[void]$menu.Items.Add($miAll)
[void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$menu.Items.Add($miTop)
[void]$menu.Items.Add($miExit)

$panel.ContextMenuStrip = $menu
$box.ContextMenuStrip   = $menu

$miNew.Add_Click({ Add-Note })
$miDel.Add_Click({ Remove-Note })
$miPrev.Add_Click({ Go-Index ($script:S.Index - 1) })
$miNext.Add_Click({ Go-Index ($script:S.Index + 1) })
$miFontUp.Add_Click({ Step-Font 1 })
$miFontDn.Add_Click({ Step-Font -1 })
$miFontRe.Add_Click({ Set-FontPx 14 })
$miSizeRe.Add_Click({
    $script:S.WW = [int][math]::Round(380 * [double]$script:S.K)
    $script:S.HH = [int][math]::Round(300 * [double]$script:S.K)
    Apply-Layout
    Save-Data
})
$miCopy.Add_Click({ try { $script:S.Box.Copy() } catch { } })
$miPaste.Add_Click({ try { $script:S.Box.Focus(); $script:S.Box.Paste() } catch { } })
$miAll.Add_Click({ try { $script:S.Box.Focus(); $script:S.Box.SelectAll() } catch { } })

$miTop.Add_Click({
    $script:S.TopMost = -not $script:S.TopMost
    $script:S.Form.TopMost = $script:S.TopMost
    $miTop.Checked = $script:S.TopMost
    Save-Data
})
$miExit.Add_Click({ if ($script:S.Form) { $script:S.Form.Close() } })

# ---------------- 定时器 ----------------
$script:SaveTimer = New-Object System.Windows.Forms.Timer
$script:SaveTimer.Interval = 700
$script:SaveTimer.Add_Tick({
    $script:SaveTimer.Stop()
    Save-Data
    $script:S.Panel.Invalidate()
})

# ---------------- 收尾 ----------------
$form.Add_FormClosing({
    $script:dragActive = $false
    try { $script:SaveTimer.Stop() } catch { }
    Sync-Current
    Save-Data
})

$form.Add_Shown({
    Trace-Log ('shown bounds=' + $script:S.Form.Bounds + ' index=' + $script:S.Index + ' count=' + $script:S.Notes.Count + ' textlen=' + ([string]$script:S.Box.Text).Length)
    try {
        $script:S.Box.SelectionStart = $script:S.Box.TextLength
        $script:S.Box.SelectionLength = 0
    } catch { }
    $script:S.Panel.Invalidate()
})

[System.Windows.Forms.Application]::add_ThreadException({
    param($s, $e)
    Write-Log ('UI: ' + $e.Exception.Message)
})

[System.Windows.Forms.Application]::Run($form)
