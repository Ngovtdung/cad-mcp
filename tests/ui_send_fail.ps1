# Send voi claude GIA (exe thoat ma 1, stderr "X"): dong do phai chua "X" va chip tep phai quay lai hang cho, N/N lan.
# Khong chay claude that, khong chay server.py, khong cham AutoCAD. Can tests\build-test.cmd truoc (tao ClaudePalette.dll).
#   powershell -NoProfile -STA -File tests\ui_send_fail.ps1 [-N 20]
param(
  [string]$Dll = "$env:TEMP\cad-mcp-test\ClaudePalette.dll",
  [string]$Fx = "$env:TEMP\cad-mcp-test\fixtures",
  [int]$N = 20
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
[AppDomain]::CurrentDomain.add_AssemblyResolve({
  param($sender, $e)
  $f = 'C:\Program Files\Autodesk\AutoCAD 2022\' + ($e.Name -split ',')[0] + '.dll'
  if (Test-Path $f) { [Reflection.Assembly]::LoadFrom($f) }
})
$stub = Join-Path ([IO.Path]::GetTempPath()) 'cad-mcp-fake-claude.exe'
Remove-Item $stub -ErrorAction SilentlyContinue
Add-Type -TypeDefinition 'public static class FakeClaude { public static int Main() { System.Console.Error.WriteLine("X"); return 1; } }' -OutputType ConsoleApplication -OutputAssembly $stub
$t = [Reflection.Assembly]::LoadFrom($Dll).GetType('ClaudeCad.ClaudeUi', $true)
$bf = [Reflection.BindingFlags]'Instance,Static,NonPublic,Public'
$t.GetField('ExeOverride', $bf).SetValue($null, $stub)
$ui = [Activator]::CreateInstance($t)
function F($name) { $t.GetField($name, $bf).GetValue($ui) }
function M($name, [object[]]$a) {
  $b = New-Object object[] $a.Length
  for ($i = 0; $i -lt $a.Length; $i++) { if ($null -ne $a[$i]) { $b[$i] = $a[$i].psobject.BaseObject } }
  $t.GetMethod($name, $bf).Invoke($ui, $b)
}
function Pump {
  $f = New-Object Windows.Threading.DispatcherFrame
  [Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background, [Action]{ $f.Continue = $false }) | Out-Null
  [Windows.Threading.Dispatcher]::PushFrame($f)
}
function WaitFor([scriptblock]$cond, [int]$sec) {
  $sw = [Diagnostics.Stopwatch]::StartNew()
  do { Pump; Start-Sleep -Milliseconds 30 } while (-not (& $cond) -and $sw.Elapsed.TotalSeconds -lt $sec)
  Pump
}
$ok = 0
try {
  for ($i = 1; $i -le $N; $i++) {
    M 'AddFiles' @(, [string[]]@("$Fx\big.txt"))
    WaitFor { (F 'loading').Count -eq 0 } 30
    $before = (F 'pending').Count
    M 'Send' @('hello')
    WaitFor { $null -eq (F 'turn') -and (F 'log').Children.Count -gt 0 -and @((F 'log').Children | Where-Object { $_ -is [Windows.Controls.TextBox] -and $_.Text -like 'L*: X' }).Count -gt 0 } 30
    $red = @((F 'log').Children | Where-Object { $_ -is [Windows.Controls.TextBox] -and $_.Text -like 'L*: X' }).Count
    $chip = (F 'pending').Count
    if ($before -eq 1 -and $red -eq 1 -and $chip -eq 1 -and $null -eq (F 'turn')) { $ok++ } else { "FAIL lan ${i}: pending truoc=$before red=$red pending sau=$chip" }
    (F 'log').Children.Clear()
    $items = @(F 'pending'); foreach ($a in $items) { M 'RemoveAttachment' @($a) }
  }
  # ---- tien trinh am: claude gia doc stdin, tra result, giu song; 2 lan Gui phai chay tren CUNG 1 tien trinh (khong session id -> khong dung SessionFile that) ----
  $echo = Join-Path ([IO.Path]::GetTempPath()) 'cad-mcp-fake-claude-echo.exe'
  $pidf = Join-Path ([IO.Path]::GetTempPath()) 'cad-mcp-fake-claude-echo.pids'
  Remove-Item $echo, $pidf -ErrorAction SilentlyContinue
  $env:FAKE_PIDFILE = $pidf
  Add-Type -TypeDefinition 'public static class EchoClaude { public static int Main() { string l; var pf = System.Environment.GetEnvironmentVariable("FAKE_PIDFILE"); while ((l = System.Console.In.ReadLine()) != null) { System.IO.File.AppendAllText(pf, System.Diagnostics.Process.GetCurrentProcess().Id + System.Environment.NewLine); System.Console.Out.WriteLine("{\"type\":\"assistant\",\"message\":{\"content\":[{\"type\":\"text\",\"text\":\"ok\"}]}}"); System.Console.Out.WriteLine("{\"type\":\"result\",\"subtype\":\"success\",\"is_error\":false,\"result\":\"ok\",\"duration_ms\":1,\"duration_api_ms\":1,\"num_turns\":1,\"usage\":{\"input_tokens\":1,\"output_tokens\":1}}"); System.Console.Out.Flush(); } return 0; } }' -OutputType ConsoleApplication -OutputAssembly $echo
  $t.GetField('ExeOverride', $bf).SetValue($null, $echo)
  (F 'log').Children.Clear()
  M 'Send' @('mot'); WaitFor { $null -eq (F 'turn') } 30
  M 'Send' @('hai'); WaitFor { $null -eq (F 'turn') } 30
  $pids = @(Get-Content $pidf -ErrorAction SilentlyContinue)
  $live = $t.GetField('live', $bf).GetValue($ui)
  if ($pids.Count -eq 2 -and $pids[0] -eq $pids[1] -and $null -ne $live) { 'warm: 2 tin nhan, cung 1 tien trinh, van song: PASS' } else { "warm FAIL: pids=$($pids -join ',') live=$($null -ne $live)"; $ok = -1 }
  $wp = Get-Process -Id ([int]$pids[0]) -ErrorAction SilentlyContinue
  M 'Shutdown' @()
  Start-Sleep -Milliseconds 500
  if ($pids.Count -gt 0 -and (Get-Process -Id ([int]$pids[0]) -ErrorAction SilentlyContinue)) { 'warm FAIL: Shutdown khong kill tien trinh'; $ok = -1 }
  Remove-Item $echo, $pidf -ErrorAction SilentlyContinue
} finally {
  M 'Shutdown' @()
  Remove-Item $stub -ErrorAction SilentlyContinue
}
"$ok/$N PASS"
if ($ok -ne $N) { exit 1 }
