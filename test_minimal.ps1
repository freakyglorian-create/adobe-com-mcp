$ErrorActionPreference = "Continue"
try {
  Write-Output "STEP1: creating COM..."
  $app = New-Object -ComObject Photoshop.Application
  Write-Output "STEP2: app = $($app.GetType().FullName)"
  Write-Output "STEP3: version = $($app.Version)"
  Write-Output "STEP4: docs count = $($app.Documents.Count)"
  $d = $app.Documents.Add(1, 1, 800, "MinimalTest")
  Write-Output "STEP5: created doc = $($d.Name)"
}
catch {
  Write-Output "ERROR: $($_.Exception.Message)"
  Write-Output "Line: $($_.InvocationInfo.ScriptLineNumber)"
}
