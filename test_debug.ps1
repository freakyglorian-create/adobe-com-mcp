$ErrorActionPreference = 'Stop'
try {
  Write-Output "DBG1: starting"
  $app = New-Object -ComObject Photoshop.Application
  Write-Output "DBG2: app created"
  if ($null -eq $app) { Write-Output "DBG2b: app is NULL!" }
  Write-Output "DBG3: version = $($app.Version)"
  Write-Output "DBG4: docs = $($app.Documents.Count)"
  $d = $app.Documents.Add(1, 1, 800, "LineTest")
  Write-Output "DBG5: doc = $($d.Name)"
}
catch {
  Write-Output "ERROR: $($_.Exception.Message)"
  Write-Output "LINE: $($_.InvocationInfo.ScriptLineNumber)"
  Write-Output "CMD: $($_.InvocationInfo.Line)"
}
