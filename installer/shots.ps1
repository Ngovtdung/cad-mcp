# Chup anh cac trang Setup (s02..s05) cho HDSD. Chay Setup o che do thu (/TESTROOT) - khong dong vao ban cai that, khong cham AutoCAD.
# Chi chup CUA SO Setup (PrintWindow), khong chup man hinh nen. Che (to mau nen) moi dong chu co duong dan o dia (X:\...).
# Moi anh PHAI duoc nguoi xem xet truoc khi dua vao HDSD.
#   powershell -NoProfile -ExecutionPolicy Bypass -File shots.ps1 [-Exe <setup.exe>] [-OutDir <thu muc anh>]
param(
    [string]$Exe = 'C:\Autodesk\dist\cad-mcp-setup-3.1.1.exe',
    [string]$OutDir = 'C:\Autodesk\cad-mcp\installer\docs\img'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class Win {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);
}
'@
$A = [Windows.Automation.AutomationElement]
$shots = 'C:\Autodesk\dist\_shots'
Remove-Item $shots -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $shots, $OutDir | Out-Null
'' | Set-Content "$shots\.cad-mcp-testroot"

function Get-Wiz {
    foreach ($w in $A::RootElement.FindAll([Windows.Automation.TreeScope]::Children, [Windows.Automation.Condition]::TrueCondition)) {
        if ($w.Current.ClassName -eq 'TWizardForm') { return $w }
    }
    return $null
}
function Get-Els($wiz, $cls) {
    $c = New-Object Windows.Automation.PropertyCondition($A::ClassNameProperty, $cls)
    return @($wiz.FindAll([Windows.Automation.TreeScope]::Descendants, $c))
}
function Click-Next($wiz) {
    $b = @(Get-Els $wiz 'TNewButton' | Sort-Object { $_.Current.BoundingRectangle.X })
    $n = if ($b.Count -ge 2) { $b[$b.Count - 2] } else { $b[0] }
    [void][Win]::SendMessage([IntPtr]$n.Current.NativeWindowHandle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)   # BM_CLICK
}
function Shot($wiz, $name) {
    $h = [IntPtr]$wiz.Current.NativeWindowHandle
    $r = New-Object Win+RECT
    [void][Win]::GetWindowRect($h, [ref]$r)
    $w = $r.R - $r.L; $ht = $r.B - $r.T
    $bmp = New-Object Drawing.Bitmap $w, $ht
    $g = [Drawing.Graphics]::FromImage($bmp)
    $dc = $g.GetHdc()
    [void][Win]::PrintWindow($h, $dc, 2)
    $g.ReleaseHdc($dc)
    $bg = $bmp.GetPixel(8, [int]($ht / 2))
    $brush = New-Object Drawing.SolidBrush $bg
    foreach ($e in @($wiz.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition))) {
        if ($e.Current.Name -match '[A-Za-z]:\\') {
            $q = $e.Current.BoundingRectangle
            $g.FillRectangle($brush, [int]($q.X - $r.L), [int]($q.Y - $r.T), [int]$q.Width, [int]$q.Height)
        }
    }
    $g.Dispose()
    $bmp.Save((Join-Path $OutDir ($name + '.png')), [Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    "chup: $name.png ($w x $ht)"
}
function Wait-For($what, $sb, [int]$sec = 60) {
    for ($i = 0; $i -lt $sec * 4; $i++) { if (& $sb) { return } Start-Sleep -Milliseconds 250 }
    throw "het gio: $what"
}

$p = Start-Process $Exe -ArgumentList "/TESTROOT=$shots", "/DIR=$shots\cad-mcp" -PassThru
try {
    Wait-For 'cua so Setup' { $script:wiz = Get-Wiz; $null -ne $script:wiz }
    Start-Sleep 1
    Shot $wiz 's02-trang-thong-tin'
    Click-Next $wiz; Start-Sleep 1.5
    $wiz = Get-Wiz
    Shot $wiz 's03-trang-rui-ro'
    $cb = @(Get-Els $wiz 'TNewCheckBox')[0]
    [void][Win]::SendMessage([IntPtr]$cb.Current.NativeWindowHandle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)   # BM_CLICK
    Start-Sleep 0.5
    Click-Next $wiz; Start-Sleep 1.5      # -> Ready
    # SendMessage chan den khi cai xong -> dung PostMessage (BM_CLICK) de chup duoc trang "Dang cai dat"
    $wiz = Get-Wiz
    $b = @(Get-Els $wiz 'TNewButton' | Sort-Object { $_.Current.BoundingRectangle.X })
    [void][Win]::PostMessage([IntPtr]$b[$b.Count - 2].Current.NativeWindowHandle, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)
    Wait-For 'trang Dang cai dat' {
        $script:wiz = Get-Wiz
        $hit = $false
        if ($null -ne $script:wiz) { foreach ($e in (Get-Els $script:wiz 'TNewStaticText')) { if ($e.Current.Name -match '^.ang c.i') { $hit = $true } } }
        $hit
    } 20
    Shot $wiz 's04-dang-cai'
    Wait-For 'cai xong' { Test-Path "$shots\cad-mcp\mcp.json" } 180
    Start-Sleep 3
    $wiz = Get-Wiz
    Shot $wiz 's05-ket-thuc'
} finally {
    Get-Process | Where-Object { $_.Name -like 'cad-mcp-setup*' } | Stop-Process -Force -ErrorAction SilentlyContinue
}
if (Test-Path "$shots\cad-mcp\unins000.exe") {
    Start-Process "$shots\cad-mcp\unins000.exe" -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART' -Wait
    for ($i = 0; $i -lt 120 -and (Get-Process | Where-Object { $_.Name -like 'unins*' }); $i++) { Start-Sleep -Milliseconds 500 }
}
Remove-Item $shots -Recurse -Force -ErrorAction SilentlyContinue
'xong'
