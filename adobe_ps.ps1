# Adobe Photoshop COM automation dispatcher for MCP server.
# Reads the operation + params from environment variables (UTF-16 safe, no file-encoding issues).
# All sizes are created with EXACT integer pixels via the gcd trick:
#   for W x H px, set ppi = gcd(W,H), inches = W/gcd, H/gcd  -> integer inches => no rounding.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Gcd($a, $b) { while ($b) { $t = $b; $b = $a % $b; $a = $t }; return [int]$a }
function J($o) { $o | ConvertTo-Json -Compress }

try {
  $app = New-Object -ComObject Photoshop.Application
  $app.Visible = $true
  try { $app.Preferences.RulerUnits = 1 } catch {}  # inches

  switch ($env:PSMCP_TOOL) {
    'create_document' {
      $W = [int]$env:PSMCP_W
      $H = [int]$env:PSMCP_H
      $nm = $env:PSMCP_NAME
      if ($W -lt 1) { $W = 1 }
      if ($H -lt 1) { $H = 1 }
      $g = Gcd $W $H
      if ($g -lt 1) { $g = 72 }
      $ppi = $g
      $wIn = $W / $g
      $hIn = $H / $g
      if ($nm) { $d = $app.Documents.Add($wIn, $hIn, $ppi, $nm) }
      else { $d = $app.Documents.Add($wIn, $hIn, $ppi) }
      J @{ result = 'ok'; name = $d.Name; width = $W; height = $H; resolution = $ppi; unit = 'pixels' }
      break
    }
    'get_active_document_info' {
      if ($app.Documents.Count -eq 0) { J @{ result = 'ok'; active = $false; open = 0 }; break }
      $d = $app.ActiveDocument
      $w = [double]$d.Width; $h = [double]$d.Height; $r = [double]$d.Resolution
      J @{
        result      = 'ok'
        name        = $d.Name
        width       = $w
        height      = $h
        resolution  = $r
        pixelWidth  = [math]::Round($w * $r)
        pixelHeight = [math]::Round($h * $r)
        mode        = $d.Mode
        open        = $app.Documents.Count
      }
      break
    }
    'close_active_document' {
      if ($app.Documents.Count -eq 0) { J @{ result = 'ok'; closed = $false; reason = 'no documents' }; break }
      $d = $app.ActiveDocument
      # Close(2) = do not save; suppress any dialogs first
      try { $app.DisplayDialogs = 1 } catch {}  # best-effort suppress
      $d.Close(2)
      J @{ result = 'ok'; closed = $true; name = $d.Name }
      break
    }
    'list_open_documents' {
      $list = @()
      for ($i = 1; $i -le $app.Documents.Count; $i++) {
        $list += @{ name = $app.Documents.Item($i).Name }
      }
      J @{ result = 'ok'; count = $app.Documents.Count; documents = $list }
      break
    }
    'do_action' {
      $a = $env:PSMCP_ACTION
      $f = $env:PSMCP_FROM
      if (-not $a) { J @{ result = 'error'; error = 'missing action name' }; break }
      if ($f) { $app.DoAction($a, $f) } else { $app.DoAction($a) }
      J @{ result = 'ok'; action = $a; from = $f }
      break
    }
    default { J @{ result = 'error'; error = "unknown tool: $env:PSMCP_TOOL" } }
  }
}
catch {
  J @{ result = 'error'; error = $_.Exception.Message }
}
