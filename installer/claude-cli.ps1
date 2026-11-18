# Cai/kiem Claude Code cho Setup "Claude AI cho AutoCAD" (spec-installer par.4.3). Khong in du lieu tai khoan.
#   -Mode install               tai claude.exe ban moi nhat, kiem chu ky Anthropic TRUOC khi chay, roi "claude.exe install latest"
#   -Mode verify -Path <exe>    ma 0 neu chu ky hop le cua Anthropic, PBC
#   -Mode status -Path <exe>    chay "<exe> auth status" toi da 15 s (het gio -> kill). Ma 0 = da dang nhap (khong tinh key trong bien moi truong), 1 = chua, 2 = het gio/loi.
#                               KHONG in gi ra (stdout/stderr cua claude bi bo, SEC-8)
param(
    [Parameter(Mandatory = $true)][ValidateSet('install', 'verify', 'status')][string]$Mode,
    [string]$Path
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$Signer = 'Anthropic, PBC'

# Bien chon cach xac thuc cua Claude Code. PHAI trung tap literal AuthPrefixes/AuthNames trong Attachments.cs (IsAuthVar);
# tests\test_spawn_sites.py so hai ben. Chi dung cho -Mode status (xoa khoi env cua tien trinh con, khong dung env cua chinh script nay).
$AuthPrefixes = @('ANTHROPIC_', 'CLAUDE_CODE_USE_', 'CLAUDE_CODE_OAUTH_')
$AuthNames = @('AWS_BEARER_TOKEN_BEDROCK', 'CLAUDE_API_KEY')
function Test-AuthVar([string]$Name) {
    foreach ($x in $AuthPrefixes) { if ($Name.StartsWith($x, [StringComparison]::OrdinalIgnoreCase)) { return $true } }
    foreach ($x in $AuthNames) { if ([string]::Equals($Name, $x, [StringComparison]::OrdinalIgnoreCase)) { return $true } }
    return $false
}

function Test-Signed([string]$File) {
    $s = Get-AuthenticodeSignature -LiteralPath $File
    if ($s.Status -ne 'Valid' -or $null -eq $s.SignerCertificate) { return $false }
    return ($s.SignerCertificate.GetNameInfo('SimpleName', $false) -eq $Signer)
}

$rc = 5
$dir = $null
try {
    if ($Mode -eq 'status') {
        $rc = 2
        if (-not [string]::IsNullOrEmpty($Path) -and [IO.Path]::IsPathRooted($Path) -and (Test-Path -LiteralPath $Path -PathType Leaf)) {
            $p = New-Object System.Diagnostics.Process
            $p.StartInfo.FileName = $Path
            $p.StartInfo.Arguments = 'auth status'
            $p.StartInfo.UseShellExecute = $false
            $p.StartInfo.CreateNoWindow = $true
            $p.StartInfo.RedirectStandardOutput = $true
            $p.StartInfo.RedirectStandardError = $true
            # key trong env (ANTHROPIC_API_KEY...) thang phien claude.ai -> "auth status" thoat 0 du chua dang nhap. Chi tien trinh con; gom khoa roi moi xoa.
            $drop = @(); foreach ($k in $p.StartInfo.EnvironmentVariables.Keys) { if (Test-AuthVar ([string]$k)) { $drop += [string]$k } }
            foreach ($k in $drop) { $p.StartInfo.EnvironmentVariables.Remove($k) }
            [void]$p.Start()
            $null = $p.StandardOutput.ReadToEndAsync()   # rut het de khong nghen ong; noi dung bi bo
            $null = $p.StandardError.ReadToEndAsync()
            if ($p.WaitForExit(15000)) { if ($p.ExitCode -eq 0) { $rc = 0 } else { $rc = 1 } }
            else { try { $p.Kill() } catch { } }
        }
    } elseif ($Mode -eq 'verify') {
        $rc = 1
        if (-not [string]::IsNullOrEmpty($Path) -and (Test-Path -LiteralPath $Path -PathType Leaf) -and (Test-Signed $Path)) { $rc = 0 }
        else { Write-Output 'claude.exe: chu ky khong phai Anthropic, PBC' }
    } else {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
        [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
        $base = 'https://downloads.claude.ai/claude-code-releases'
        $ver = ([string](Invoke-RestMethod -Uri ($base + '/latest') -TimeoutSec 30 -UseBasicParsing)).Trim()
        if ($ver -notmatch '^\d+\.\d+\.\d+(-[\w.]+)?$') {
            Write-Output 'phien ban tra ve khong hop le'
            $rc = 2
        } else {
            $dir = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString())
            New-Item -ItemType Directory -Path $dir | Out-Null
            $f = Join-Path $dir 'claude.exe'
            Invoke-WebRequest -Uri ($base + '/' + $ver + '/win32-x64/claude.exe') -OutFile $f -TimeoutSec 600 -UseBasicParsing
            if (-not (Test-Signed $f)) {
                Write-Output 'chu ky claude.exe tai ve khong hop le - KHONG chay'
                $rc = 3
            } else {
                Write-Output ('Dang cai Claude Code ' + $ver)
                & $f install latest | Out-Null
                if ($LASTEXITCODE -ne 0) { Write-Output ('claude install loi ' + $LASTEXITCODE); $rc = 4 } else { $rc = 0 }
            }
        }
    }
} catch {
    if ($Mode -ne 'status') { Write-Output ('loi: ' + $_.Exception.Message) }
    $rc = 5
    if ($Mode -eq 'status') { $rc = 2 }
} finally {
    if ($dir) { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
}
exit $rc
