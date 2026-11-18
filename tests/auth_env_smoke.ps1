# auth_env_smoke.ps1 - spec-3.1.0 par.9 S2 + S3. Windows PowerShell 5.1 (cung StringDictionary voi DLL net46).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\auth_env_smoke.ps1 [-NoAuthEnv]
# Dung env kieu Explorer (Machine + User doc tu registry, CHI DOC), goi ClaudeCad.Attach.AuthEnv tu DLL net46 vua build,
# chay "claude auth status --json" voi env nguyen trang (C1) va env da loc (C2). Khong sua env that, khong ghi tep, khong goi "claude -p".
# Chi in: loggedIn, authMethod, apiKeySource, TEN bien, ma thoat, PASS/FAIL/SKIP. Ma thoat 0 = PASS/SKIP, 1 = FAIL.
# -NoAuthEnv = bo qua AuthEnv (kiem chung ban than test: PHAI FAIL).
# Buoc S3: chay installer\claude-cli.ps1 -Mode status trong env co key gia + CLAUDE_CONFIG_DIR rong -> PHAI ra ma 1.
param([switch]$NoAuthEnv, [switch]$Child)
$ErrorActionPreference = 'Stop'

# DLL bi khoa toi khi tien trinh thoat -> phan nap DLL chay trong mot PowerShell con.
if (-not $Child) {
    $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $a = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Child')
    if ($NoAuthEnv) { $a += '-NoAuthEnv' }
    & $ps @a
    exit $LASTEXITCODE
}

function Say([string]$m) { [Console]::WriteLine($m) }
function Safe($v) { $t = [string]$v; if ($t -match '^[A-Za-z0-9_.\-]{0,40}$') { return $t } else { return '?' } }

$dummy = 'dummy-not-a-key'
$exe = Join-Path $env:USERPROFILE '.local\bin\claude.exe'
$dll = Join-Path $PSScriptRoot '..\installer\bin\Release\net46\ClaudePalette.dll'
$cli = Join-Path $PSScriptRoot '..\installer\claude-cli.ps1'
$ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'

if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { Say 'SKIP: khong co claude.exe'; exit 0 }

function Read-RegEnv($hive, [string]$path) {
    $h = @{}
    $key = $hive.OpenSubKey($path)          # chi doc
    if ($key) { foreach ($n in $key.GetValueNames()) { if ($n) { $h[$n] = [string]$key.GetValue($n) } }; $key.Close() }
    return $h
}

# Env kieu Explorer: bien phien + Machine + User, Path = Machine;User
function Build-Env {
    $m = Read-RegEnv ([Microsoft.Win32.Registry]::LocalMachine) 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
    $u = Read-RegEnv ([Microsoft.Win32.Registry]::CurrentUser) 'Environment'
    $e = @{}
    foreach ($n in 'SystemRoot', 'SystemDrive', 'windir', 'ProgramData', 'ProgramFiles', 'ProgramFiles(x86)', 'ProgramW6432', 'CommonProgramFiles', 'CommonProgramFiles(x86)', 'CommonProgramW6432', 'ALLUSERSPROFILE', 'PUBLIC', 'USERPROFILE', 'USERNAME', 'USERDOMAIN', 'COMPUTERNAME', 'HOMEDRIVE', 'HOMEPATH', 'APPDATA', 'LOCALAPPDATA', 'LOGONSERVER', 'SESSIONNAME', 'NUMBER_OF_PROCESSORS', 'PROCESSOR_ARCHITECTURE', 'OS', 'ComSpec', 'PATHEXT', 'TEMP', 'TMP') {
        $v = [Environment]::GetEnvironmentVariable($n, 'Process'); if ($v) { $e[$n] = $v }
    }
    foreach ($kv in $m.GetEnumerator()) { if ($kv.Key -ne 'Path') { $e[$kv.Key] = $kv.Value } }
    foreach ($kv in $u.GetEnumerator()) { if ($kv.Key -ne 'Path') { $e[$kv.Key] = $kv.Value } }
    $e['Path'] = [string]$m['Path'] + ';' + [string]$u['Path']
    return $e
}

function New-Psi([string]$File, [string]$ArgLine, $EnvMap) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo($File, $ArgLine)
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $psi.WorkingDirectory = [IO.Path]::GetTempPath()
    $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = [Text.Encoding]::UTF8
    $psi.EnvironmentVariables.Clear()
    foreach ($k in $EnvMap.Keys) { $psi.EnvironmentVariables[[string]$k] = [string]$EnvMap[$k] }
    return $psi
}

# Tra @{Code; Out}; Code = $null khi het gio (da Kill). stderr bi bo.
function Invoke-Psi($psi, [int]$TimeoutMs) {
    $p = [Diagnostics.Process]::Start($psi)
    $o = $p.StandardOutput.ReadToEndAsync(); $null = $p.StandardError.ReadToEndAsync()
    if (-not $p.WaitForExit($TimeoutMs)) { try { $p.Kill() } catch { }; $p.Dispose(); return @{ Code = $null; Out = '' } }
    $p.WaitForExit()
    $r = @{ Code = $p.ExitCode; Out = $o.Result }
    $p.Dispose()
    return $r
}

function Get-Status($r) {
    $s = @{ Ok = $false; LoggedIn = ''; Method = ''; Source = '' }
    if ($null -eq $r.Code) { return $s }
    try {
        $j = $r.Out | ConvertFrom-Json
        $s.LoggedIn = Safe $j.loggedIn; $s.Method = Safe $j.authMethod; $s.Source = Safe $j.apiKeySource
        $s.Ok = $true
    } catch { }
    return $s
}

function Show([string]$label, $s, $code) {
    Say ($label + ': exit=' + $code + ' loggedIn=' + $s.LoggedIn + ' authMethod=' + $s.Method + ' apiKeySource=' + $s.Source)
}

function Names($sd, $other) {   # ten (HOA) co trong $sd ma khong co trong $other
    $o = @{}; foreach ($k in $other.Keys) { $o[[string]$k] = 1 }
    return ((@($sd.Keys | Where-Object { -not $o.ContainsKey([string]$_) } | ForEach-Object { ([string]$_).ToUpperInvariant() }) | Sort-Object) -join ',')
}

$fail = New-Object System.Collections.ArrayList
$tmpCfg = $null
try {
    # ---- nap AuthEnv that tu DLL; KHONG co duong du phong ----
    $mi = $null
    try {
        $asm = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $dll).ProviderPath)
        $mi = $asm.GetType('ClaudeCad.Attach').GetMethod('AuthEnv')
    } catch { Say ('DLL: khong nap duoc (' + $_.Exception.GetType().Name + ')') }
    if ($asm) { Say ('DLL: ClaudePalette version ' + $asm.GetName().Version.ToString()) }
    if ($null -eq $mi) { [void]$fail.Add('khong nap duoc ClaudeCad.Attach.AuthEnv tu DLL net46') }
    else {
        $clean = Build-Env
        $injected = -not $clean.ContainsKey('ANTHROPIC_API_KEY')
        if ($injected) { $clean['ANTHROPIC_API_KEY'] = $dummy }
        Say ('ANTHROPIC_API_KEY gia duoc chen: ' + $(if ($injected) { 'co' } else { 'khong (may co san)' }))

        # C1: env nguyen trang
        $psi1 = New-Psi $exe 'auth status --json' $clean
        $r1 = Invoke-Psi $psi1 30000; $s1 = Get-Status $r1; Show 'C1 nguyen trang' $s1 $r1.Code

        # C2: env da loc bang AuthEnv that (key = null, object[] de PS khong doi null thanh "")
        $psi2 = New-Psi $exe 'auth status --json' $clean
        $before = New-Object System.Collections.Specialized.StringDictionary
        foreach ($k in $psi2.EnvironmentVariables.Keys) { $before[$k] = $psi2.EnvironmentVariables[$k] }
        if (-not $NoAuthEnv) { $ar = New-Object object[] 2; $ar[0] = $psi2.EnvironmentVariables; $ar[1] = $null; [void]$mi.Invoke($null, $ar) }
        else { Say 'C2: -NoAuthEnv, bo qua AuthEnv' }
        Say ('C2 ten bien da xoa: ' + (Names $before $psi2.EnvironmentVariables))
        # con khoa nao thoa IsAuthVar? Goi lai AuthEnv tren ban sao: phai khong xoa them gi.
        $copy = New-Object System.Collections.Specialized.StringDictionary
        foreach ($k in $psi2.EnvironmentVariables.Keys) { $copy[$k] = $psi2.EnvironmentVariables[$k] }
        $ar = New-Object object[] 2; $ar[0] = [Collections.Specialized.StringDictionary]$copy; $ar[1] = $null; [void]$mi.Invoke($null, $ar)   # cast: New-Object boc PSObject
        $left = Names $psi2.EnvironmentVariables $copy
        Say ('C2 ten bien auth con lai: ' + $(if ($left) { $left } else { '(khong)' }))
        $r2 = Invoke-Psi $psi2 30000; $s2 = Get-Status $r2; Show 'C2 da loc' $s2 $r2.Code

        if (-not $s1.Ok -or $s1.Method -ne 'api_key') { [void]$fail.Add('tien dieu kien: env nguyen trang khong ra api_key') }
        if (-not $s2.Ok) { [void]$fail.Add('C2: khong doc duoc auth status') }
        if ($s2.Method -eq 'api_key') { [void]$fail.Add('C2: env da loc van ra api_key') }
        if ($s2.Source) { [void]$fail.Add('C2: apiKeySource khong rong') }
        if ($left) { [void]$fail.Add('C2: con bien thoa IsAuthVar') }

        # S3: bo cai (-Mode status) trong env co key gia + CLAUDE_CONFIG_DIR rong -> chua dang nhap -> ma 1
        if (-not (Test-Path -LiteralPath $cli -PathType Leaf)) { [void]$fail.Add('S3: khong thay claude-cli.ps1') }
        else {
            $tmpCfg = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())
            [void](New-Item -ItemType Directory -Path $tmpCfg)
            $e3 = @{}; foreach ($k in $clean.Keys) { $e3[$k] = $clean[$k] }
            $e3['ANTHROPIC_API_KEY'] = $dummy; $e3['CLAUDE_CONFIG_DIR'] = $tmpCfg
            $argLine = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + (Resolve-Path -LiteralPath $cli).ProviderPath + '" -Mode status -Path "' + $exe + '"'
            $r3 = Invoke-Psi (New-Psi $ps $argLine $e3) 60000
            Say ('S3 claude-cli.ps1 -Mode status (key gia, config rong): exit=' + $r3.Code + ' (mong doi 1)')
            if ($r3.Code -ne 1) { [void]$fail.Add('S3: ma thoat khac 1') }
        }
    }
} catch { [void]$fail.Add('loi test: ' + $_.Exception.GetType().Name) }
finally { if ($tmpCfg) { Remove-Item -LiteralPath $tmpCfg -Recurse -Force -ErrorAction SilentlyContinue } }

if ($fail.Count -eq 0) { Say 'PASS'; exit 0 }
foreach ($f in $fail) { Say ('FAIL: ' + $f) }
exit 1
