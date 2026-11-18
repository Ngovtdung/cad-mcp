# Kiem Authenticode. Ma 0 khi dat, hong: in 1 dong ly do, ma 1.
#   -Signer <ten>        Status=Valid va SimpleName nguoi ky = <ten> (nghiem)
#   -MicrosoftToolchain  Status=Valid, Subject chua "O=Microsoft Corporation" va Issuer = "CN=Microsoft Code Signing PCA *, O=Microsoft Corporation"
#                        (chi dung cho dotnet.exe: SimpleName cua no la ".NET")
param([Parameter(Mandatory = $true)][string]$Path, [string]$Signer, [switch]$MicrosoftToolchain)
if ((-not $MicrosoftToolchain -and -not $Signer) -or ($MicrosoftToolchain -and $Signer)) { Write-Output 'Dung dung mot trong -Signer / -MicrosoftToolchain'; exit 2 }
$s = Get-AuthenticodeSignature -LiteralPath $Path
$c = $s.SignerCertificate
$n = ''; $sub = ''; $iss = ''
if ($null -ne $c) { $n = $c.GetNameInfo('SimpleName', $false); $sub = $c.Subject; $iss = $c.Issuer }
if ($MicrosoftToolchain) {
    if ($s.Status -eq 'Valid' -and $sub -match '(^|, )O=Microsoft Corporation(,|$)' -and $iss -match '^CN=Microsoft Code Signing PCA [^,]+, O=Microsoft Corporation(,|$)') { exit 0 }
} elseif ($s.Status -eq 'Valid' -and $n -eq $Signer) { exit 0 }
Write-Output ('CHU KY HONG: ' + $Path + ' | ' + $s.Status + ' | ' + $n)
exit 1