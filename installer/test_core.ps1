# Kiem ClaudeCad.Core (IsDriveAbs, ResolveHome, FindClaude, ParseJson) trong 1 DLL - khong can AutoCAD.
#   powershell -NoProfile -File test_core.ps1 -Dll ...\net46\ClaudePalette.dll         (.NET Framework)
#   pwsh -NoProfile -File test_core.ps1 -Dll ...\net8.0-windows\ClaudePalette.dll       (.NET)
# 3.1: net10.0-windows cũng nạp được bằng pwsh; máy chưa có pwsh thì net8/net10 chỉ kiểm bằng build + ma trận spec-3.1.0 §5.
param([Parameter(Mandatory = $true)][string]$Dll)
$ErrorActionPreference = 'Stop'
$t = [Reflection.Assembly]::LoadFrom((Resolve-Path $Dll).Path).GetType('ClaudeCad.Core', $true)
function Call($name, [object[]]$a) { $u = [object[]]@($a | ForEach-Object { if ($_ -is [psobject]) { $_.psobject.BaseObject } else { $_ } }); $t.GetMethod($name).Invoke($null, $u) }  # Invoke không nhận PSObject bọc
$script:n = 0; $script:fail = 0
function Check($label, $got, $want) {
  $script:n++
  if ($got -ne $want) { $script:fail++; Write-Host "FAIL $label`n  got:  $got`n  want: $want" }
}

$tmp = Join-Path ([IO.Path]::GetTempPath()) ('core-test-' + [guid]::NewGuid())
$lad = Join-Path $tmp 'lad'; $homeOk = Join-Path $tmp 'home ok'; $homeNo = Join-Path $tmp 'home-no-mcp'
$prof = Join-Path $tmp 'profile'; $dirA = Join-Path $tmp 'a'; $dirB = Join-Path $tmp 'b c'
New-Item -ItemType Directory $lad, $homeOk, $homeNo, "$prof\.local\bin", $dirA, $dirB | Out-Null
'{}' | Set-Content "$homeOk\mcp.json"
'x' | Set-Content "$dirA\claude.exe"; 'x' | Set-Content "$dirB\claude.exe"
try {
  Check 'home từ registry' (Call ResolveHome @($homeOk, $lad)) $homeOk
  Check 'home thiếu mcp.json' (Call ResolveHome @($homeNo, $lad)) (Join-Path $lad 'cad-mcp')
  Check 'home null' (Call ResolveHome @($null, $lad)) (Join-Path $lad 'cad-mcp')
  Check 'home tương đối' (Call ResolveHome @('home ok', $lad)) (Join-Path $lad 'cad-mcp')
  Check 'home UNC bị loại' (Call ResolveHome @('\\srv\share\cad-mcp', $lad)) (Join-Path $lad 'cad-mcp')
  Push-Location $tmp   # thư mục hiện hành có a\claude.exe -> mục PATH tương đối 'a' phải bị bỏ qua
  try {
    Check 'PATH bỏ mục tương đối, bỏ dấu nháy' (Call FindClaude @($prof, "a;`"$dirB`"")) (Join-Path $dirB 'claude.exe')
    Check 'PATH ký tự lạ không ném lỗi' (Call FindClaude @($prof, "C:\x|y;$dirA")) (Join-Path $dirA 'claude.exe')
    Check 'PATH UNC và \x bị loại' (Call FindClaude @($prof, '\\srv\share;\x')) $null
    Check 'không có claude' (Call FindClaude @($prof, '')) $null
    'x' | Set-Content "$prof\.local\bin\claude.exe"
    Check 'ưu tiên ~\.local\bin' (Call FindClaude @($prof, $dirA)) (Join-Path $prof '.local\bin\claude.exe')
  } finally { Pop-Location }
  $d = Call ParseJson @('{"loggedIn":true,"email":"a@b.c","n":3,"message":{"content":[{"type":"text","text":"xin chào"}]}}')
  Check 'json bool' (([bool]$true).Equals($d['loggedIn'])) $true
  Check 'json chuỗi' $d['email'] 'a@b.c'
  $c = @($d['message']['content'])
  Check 'json mảng chứa dict' $c[0]['text'] 'xin chào'
  Check 'json hỏng ném lỗi' $(try { Call ParseJson @('{x') | Out-Null; 'không ném' } catch { 'ném' }) 'ném'
} finally { Remove-Item -Recurse -Force $tmp }
"$($script:n - $script:fail)/$($script:n) PASS"
if ($script:fail) { exit 1 }