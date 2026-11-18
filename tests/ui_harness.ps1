# Lái UI palette ngoài AutoCAD: nạp DLL bằng reflection, đặt Page vào Window, gọi hàm đính kèm, chụp PNG.
#   powershell -NoProfile -STA -File tests\ui_harness.ps1 [-Answer <tệp chữ: câu trả lời thật của claude>]
# KHÔNG gọi Send (chạy claude + server.py -> chạm AutoCAD), NewChat (xoá tệp session thật), CaptureRegion. KHÔNG đụng clipboard thật.
# 'sendlog.png' = bong bóng user dựng bằng UserBubble + câu trả lời claude thật (-Answer) đưa vào qua OnEvent.
param(
  [string]$Dll = "$env:TEMP\cad-mcp-test\ClaudePalette.dll",
  [string]$Fx = "$env:TEMP\cad-mcp-test\fixtures",
  [string]$FxXls = "$PSScriptRoot\fixtures-xls",
  [string]$Out = "$env:TEMP\cad-mcp-test\ui",
  [string]$Answer = ""
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
New-Item -ItemType Directory -Force $Out | Out-Null
# ngoài AutoCAD: JIT các hàm của ClaudeUi cần NẠP (không chạy) accoremgd/acdbmgd để phân giải kiểu -> trỏ về thư mục AutoCAD (chỉ đọc DLL)
[AppDomain]::CurrentDomain.add_AssemblyResolve({
  param($sender, $e)
  $f = 'C:\Program Files\Autodesk\AutoCAD 2022\' + ($e.Name -split ',')[0] + '.dll'
  if (Test-Path $f) { [Reflection.Assembly]::LoadFrom($f) }
})
$t = [Reflection.Assembly]::LoadFrom($Dll).GetType('ClaudeCad.ClaudeUi', $true)
$ui = [Activator]::CreateInstance($t)
$page = $t.GetField('Page').GetValue($ui)
$bf = [Reflection.BindingFlags]'Instance,Static,NonPublic,Public'
function M($name, [object[]]$a) {
  $b = New-Object object[] $a.Length   # MethodInfo.Invoke không tự bỏ lớp PSObject bọc ngoài đối số
  for ($i = 0; $i -lt $a.Length; $i++) { if ($null -ne $a[$i]) { $b[$i] = $a[$i].psobject.BaseObject } }
  $t.GetMethod($name, $bf).Invoke($ui, $b)
}
function F($name) { $t.GetField($name, $bf).GetValue($ui) }
function Pump {
  $f = New-Object Windows.Threading.DispatcherFrame
  [Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background, [Action]{ $f.Continue = $false }) | Out-Null
  [Windows.Threading.Dispatcher]::PushFrame($f)
}
function WaitIdle {   # đợi luồng nền đọc xong (Finish chạy qua Dispatcher của luồng này)
  $sw = [Diagnostics.Stopwatch]::StartNew()
  do { Pump; Start-Sleep -Milliseconds 30 } while ((F 'loading').Count -gt 0 -and $sw.Elapsed.TotalSeconds -lt 60)
  Pump
}
function Shot($name) {
  Pump; $page.Measure((New-Object Windows.Size 380, 760)); $page.Arrange((New-Object Windows.Rect 0, 0, 380, 760)); $page.UpdateLayout()
  $rtb = New-Object Windows.Media.Imaging.RenderTargetBitmap ([int]$page.ActualWidth), ([int]$page.ActualHeight), 96, 96, ([Windows.Media.PixelFormats]::Pbgra32)
  $rtb.Render($page)
  $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
  $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
  $fs = [IO.File]::Create("$Out\$name.png"); $enc.Save($fs); $fs.Close()
  "SHOT $Out\$name.png"
}
function Names { ((F 'pending') | ForEach-Object { $_.Name }) -join ', ' }
function LogLines {
  foreach ($c in (F 'log').Children) {
    if ($c -is [Windows.Controls.TextBox]) { 'LOG: ' + $c.Text }
    elseif ($c -is [Windows.Controls.TextBlock]) { 'NOTE: ' + $c.Text }
    else { 'ELEM: ' + $c.GetType().Name }
  }
}
function Data($format, $value) { $d = New-Object Windows.DataObject; $d.SetData($format, $value); $d }

# KHÔNG dùng Window: Window.Show làm ô nhập nhận focus -> GotKeyboardFocus -> Commands.SetKeepFocus cần acdbmgd (C++/CLI, chỉ nạp được trong AutoCAD).
# Page không gắn PresentationSource thì Focus() không làm gì; tự đo/xếp bố cục để chụp ảnh.
$page.Measure((New-Object Windows.Size 380, 760)); $page.Arrange((New-Object Windows.Rect 0, 0, 380, 760)); $page.UpdateLayout()
Shot 'empty'

M 'AddFiles' @(, [string[]](1..8 | ForEach-Object { "$Fx\bad$_.pdf" }))
WaitIdle
"A0 pending=$((F 'pending').Count)"
Shot 'errors'   # 5 dòng đỏ + 1 dòng gộp trong log

M 'AddFiles' @(, [string[]]@("$Fx\big.png", "$Fx\doc.pdf", "$Fx\book.xlsx", "$Fx\vi_utf8.txt", "$Fx\big.txt", "$Fx\old.xls", $Fx))
"A1 loading=$((F 'loading').Count) pending=$((F 'pending').Count) kind0=$((F 'pending')[0].Kind)"
Shot 'loading'
WaitIdle
"A2 pending=$((F 'pending').Count) [$(Names)] chipsVisible=$((F 'chips').Visibility) chipCount=$((F 'chips').Children.Count)"
Shot 'chips'

# đủ mọi loại chip: ảnh, PDF, xlsx, xls thật, văn bản, văn bản bị cắt (đỏ). Rồi bỏ chip xls để các số đếm sau giữ nguyên
if (Test-Path "$FxXls\book.xls") {
  M 'AddFiles' @(, [string[]]@("$FxXls\book.xls"))
  WaitIdle
  "A3 pending=$((F 'pending').Count) [$(Names)]"
  Shot 'chips-all'
  M 'RemoveAttachment' @((F 'pending')[(F 'pending').Count - 1])
} else { "A3 SKIP chưa có book.xls" }

M 'ShowDropHint' @($true); "B dropHint=$((F 'dropHint').Visibility)"; Shot 'drop'; M 'ShowDropHint' @($false)

"C drop=$(M 'SafeTake' @((Data ([Windows.DataFormats]::FileDrop) ([string[]]@("$Fx\small.png"))), $false))"; WaitIdle; "C pending=$((F 'pending').Count)"
"D textPaste=$(M 'TakeData' @((Data ([Windows.DataFormats]::UnicodeText) 'abc'), $true)) pending=$((F 'pending').Count)"
"E textDrop=$(M 'TakeData' @((Data ([Windows.DataFormats]::UnicodeText) 'abc'), $false)) hasFiles=$($t.GetMethod('HasFiles', $bf).Invoke($null, @((Data ([Windows.DataFormats]::UnicodeText) 'abc').psobject.BaseObject)))"
$bs = [Windows.Media.Imaging.BitmapSource]::Create(50, 40, 96, 96, [Windows.Media.PixelFormats]::Bgra32, $null, (New-Object byte[] (50 * 40 * 4)), 200)
"F bitmapPaste=$(M 'TakeData' @((Data ([Windows.DataFormats]::Bitmap) $bs), $true))"; WaitIdle; "F pending=$((F 'pending').Count) [$(Names)]"

$x = (F 'chips').Children[0].Child.Children[0]
$ev = New-Object Windows.Input.MouseButtonEventArgs ([Windows.Input.Mouse]::PrimaryDevice), 0, ([Windows.Input.MouseButton]::Left)
$ev.RoutedEvent = [Windows.UIElement]::MouseLeftButtonUpEvent
$x.RaiseEvent($ev)
"G afterRemove pending=$((F 'pending').Count) chipCount=$((F 'chips').Children.Count) [$(Names)]"

M 'AddFiles' @(, [string[]]@("$Fx\meta.png"))
$ph = (F 'pending')[(F 'pending').Count - 1]
"H0 kind=$($ph.Kind) name=$($ph.Name)"
M 'RemoveAttachment' @($ph)
WaitIdle
"H pending=$((F 'pending').Count) [$(Names)]"

M 'AddFiles' @(, [string[]]@("$Fx\doc.pdf", "$Fx\vi_utf8bom.lsp", "$Fx\cafe_1252.txt", "$Fx\vi_1258.csv", "$Fx\real.xlsx"))
WaitIdle
"I pending=$((F 'pending').Count) [$(Names)]"
M 'AddFiles' @(, [string[]]@("$Fx\vi_utf16le.txt"))
WaitIdle
"J pending=$((F 'pending').Count)"

$pend = $t.GetField('pending', $bf).GetValue($ui)   # gán thẳng: qua hàm F thì List bị trải ra thành object[]
M 'Append' @(, (M 'UserBubble' @('Vẽ theo ảnh này', $pend)))
Shot 'bubble'
LogLines

# nhật ký sau khi gửi: ảnh + PDF + xlsx + xls trong bong bóng user, rồi câu trả lời thật của claude (-Answer)
foreach ($a in $pend.ToArray()) { M 'RemoveAttachment' @($a) }
$xlsFiles = @("$Fx\big.png", "$Fx\doc.pdf", "$Fx\book.xlsx")
if (Test-Path "$FxXls\book.xls") { $xlsFiles += "$FxXls\book.xls" }
M 'AddFiles' @(, [string[]]$xlsFiles)
WaitIdle
"K pending=$((F 'pending').Count) [$(Names)]"
$sent = [Collections.Generic.List[ClaudeCad.Attachment]]::new()
foreach ($a in (F 'pending')) { $sent.Add($a) }
(F 'log').Children.Clear()
M 'Append' @(, (M 'UserBubble' @('Tóm tắt các tệp đính kèm', $sent)))
if ($Answer -ne "" -and (Test-Path $Answer)) { M 'OnEvent' @('text', ([IO.File]::ReadAllText($Answer, [Text.Encoding]::UTF8).Trim())) }
Shot 'sendlog'
LogLines

# ---------- phím tắt (spec-hotkeys §8). Sao lưu/khôi phục registry thật; không Hotkeys.Start -> không đăng ký phím thật ----------
$rk = 'HKCU:\Software\ClaudeCad'
$bak = @{}
foreach ($n in 'HotkeyCapture', 'HotkeyToggle') { $bak[$n] = (Get-ItemProperty $rk -Name $n -ErrorAction SilentlyContinue).$n }
$hk = $t.Assembly.GetType('ClaudeCad.Hotkeys', $true)
$CS = [Windows.Input.ModifierKeys]::Control -bor [Windows.Input.ModifierKeys]::Shift   # cast chuỗi 'Control,Shift' bị ModifierKeysConverter của WPF từ chối
try {
  Remove-ItemProperty $rk -Name HotkeyCapture, HotkeyToggle -ErrorAction SilentlyContinue
  M 'ApplyHotkeys' @(); M 'ShowHotkeys' @($false)
  "HK0 capture=$($page.FindName('HotkeyCapture').Text) toggle=$($page.FindName('HotkeyToggle').Text)"   # mặc định
  $r = M 'ApplyHotkeyInput' @('HotkeyToggle', [Windows.Input.Key]::D7, ($CS))
  "HK1 ok=$r reg=$((Get-ItemProperty $rk).HotkeyToggle) box=$($page.FindName('HotkeyToggle').Text)"
  # VK 0 (vd VK_PACKET của UniKey -> Key.None) và Backspace không được tắt phím; registry giữ nguyên
  $r = M 'ApplyHotkeyInput' @('HotkeyToggle', [Windows.Input.Key]::None, [Windows.Input.ModifierKeys]::Control)
  "HK1b Key.None ok=$r reg=$((Get-ItemProperty $rk).HotkeyToggle)"
  $r = M 'ApplyHotkeyInput' @('HotkeyToggle', [Windows.Input.Key]::Back, [Windows.Input.ModifierKeys]::None)
  "HK1c Backspace ok=$r reg=$((Get-ItemProperty $rk).HotkeyToggle)"
  $r = M 'ApplyHotkeyInput' @('HotkeyCapture', [Windows.Input.Key]::None, [Windows.Input.ModifierKeys]::Control)
  "HK1d capture Key.None ok=$r regExists=$($null -ne (Get-ItemProperty $rk).HotkeyCapture)"
  $page.FindName('SettingsSection').IsExpanded = $true
  Shot 'hotkeys-ok'   # trạng thái bình thường: Ctrl+Shift+X / Ctrl+Shift+7
  $r = M 'ApplyHotkeyInput' @('HotkeyToggle', [Windows.Input.Key]::C, ([Windows.Input.ModifierKeys]::Control))
  "HK2 ok=$r reg=$((Get-ItemProperty $rk).HotkeyToggle)"; LogLines | Select-Object -Last 1
  $r = M 'ApplyHotkeyInput' @('HotkeyToggle', [Windows.Input.Key]::X, ($CS))
  "HK3 ok=$r"; LogLines | Select-Object -Last 1
  $r = M 'ApplyHotkeyInput' @('HotkeyToggle', [Windows.Input.Key]::LeftCtrl, ([Windows.Input.ModifierKeys]::Control))
  "HK4 ok=$r (chỉ bổ trợ: chờ tiếp)"
  M 'ClearHotkey' @('HotkeyToggle')
  "HK5 reg='$((Get-ItemProperty $rk).HotkeyToggle)' box=$($page.FindName('HotkeyToggle').Text)"
  $hk.GetField('CaptureError').SetValue($null, 'Ctrl+Shift+X đang được chương trình khác dùng (thử harness)')
  M 'ShowHotkeys' @($true)
  "HK6 box=$($page.FindName('HotkeyCapture').Text) tip=$($page.FindName('HotkeyCapture').ToolTip)"; LogLines | Select-Object -Last 1
  $page.FindName('SettingsSection').IsExpanded = $true
  Shot 'hotkeys'
} finally {
  foreach ($n in $bak.Keys) {
    if ($null -eq $bak[$n]) { Remove-ItemProperty $rk -Name $n -ErrorAction SilentlyContinue } else { Set-ItemProperty $rk -Name $n -Value $bak[$n] }
  }
}
