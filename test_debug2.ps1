$ErrorActionPreference = 'Continue'
$app = New-Object -ComObject Photoshop.Application
Write-Output "=== Type ==="
$app.GetType().FullName
Write-Output "=== Members (first 30) ==="
$app | Get-Member | Select-Object -First 30 Name, MemberType | Format-Table -AutoSize
Write-Output "=== Docs COM object type ==="
$app.Documents.GetType().FullName
Write-Output "=== Docs Count ==="
$app.Documents.Count
Write-Output "=== Try Add with fewer args ==="
try {
  $d = $app.Documents.Add()
  Write-Output "SUCCESS: $($d.Name)"
} catch {
  Write-Output "FAIL Add(): $($_.Exception.Message)"
}
