# Photoshop COM automation dispatcher (extended).
# Exact integer-pixel sizing via gcd trick.
# Built-in retry for RPC_E_CALL_REJECTED (0x8001010A) — common when PS is busy.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Gcd($a, $b) { while ($b) { $t = $b; $b = $a % $b; $a = $t }; return [int]$a }
function J($o) { $o | ConvertTo-Json -Compress -Depth 4 }

function NewRgbColor($r, $g, $b) {
  $c = New-Object -ComObject Photoshop.SolidColor
  $c.RGB.Red = [double]$r
  $c.RGB.Green = [double]$g
  $c.RGB.Blue = [double]$b
  return $c
}

# Get PS Application object with retry (handles RPC_E_CALL_REJECTED)
function Get-PSApp {
  param([int]$MaxTries = 20, [int]$SleepMs = 250)
  for ($i = 1; $i -le $MaxTries; $i++) {
    try {
      $app = New-Object -ComObject Photoshop.Application
      try { $app.DisplayDialogs = 1 } catch {}
      try { $app.Preferences.RulerUnits = 1 } catch {}
      return $app
    } catch {
      if ($i -eq $MaxTries) { throw }
      Start-Sleep -Milliseconds $SleepMs
    }
  }
}

try {
  $app = Get-PSApp

  switch ($env:PSMCP_TOOL) {

    # ------ document basics ------
    'create_document' {
      $W = [int]$env:PSMCP_W; $H = [int]$env:PSMCP_H
      $nm = $env:PSMCP_NAME
      if ($W -lt 1) { $W = 800 }; if ($H -lt 1) { $H = 800 }
      $g = Gcd $W $H; if ($g -lt 1) { $g = 72 }
      if ($nm) { $d = $app.Documents.Add($W/$g, $H/$g, $g, $nm) }
      else { $d = $app.Documents.Add($W/$g, $H/$g, $g) }
      J @{ result='ok'; name=$d.Name; width=$W; height=$H; unit='pixels' }
      break
    }
    'get_active_document_info' {
      if ($app.Documents.Count -eq 0) { J @{result='ok'; active=$false; open=0}; break }
      $d = $app.ActiveDocument
      $w=[double]$d.Width; $h=[double]$d.Height; $r=[double]$d.Resolution
      J @{ result='ok'; name=$d.Name
           pixelWidth=[math]::Round($w*$r); pixelHeight=[math]::Round($h*$r)
           resolution=$r; mode=$d.Mode; open=$app.Documents.Count
           layers=$d.Layers.Count; activeLayer=$d.ActiveLayer.Name }
      break
    }
    'list_open_documents' {
      $list = @()
      for ($i = 1; $i -le $app.Documents.Count; $i++) { $list += @{ name=$app.Documents.Item($i).Name } }
      J @{ result='ok'; count=$app.Documents.Count; documents=$list }
      break
    }
    'close_active_document' {
      if ($app.Documents.Count -eq 0) { J @{result='ok'; closed=$false}; break }
      $nm = $app.ActiveDocument.Name
      $app.ActiveDocument.Close(2)
      J @{ result='ok'; closed=$true; name=$nm }
      break
    }
    'open_document' {
      $p = $env:PSMCP_PATH
      if (-not $p -or -not (Test-Path $p)) { J @{result='error'; error="file not found: $p"}; break }
      $d = $app.Open($p)
      J @{ result='ok'; name=$d.Name; path=$p }
      break
    }
    'save_as_psd' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $p = $env:PSMCP_PATH
      if (-not $p) { $p = Join-Path $env:USERPROFILE "Documents\$($app.ActiveDocument.Name).psd" }
      $psdOpts = New-Object -ComObject Photoshop.PhotoshopSaveOptions
      $psdOpts.EmbedColorProfile = $true
      $psdOpts.AlphaChannels = $true
      $psdOpts.Layers = $true
      $psdOpts.Spots = $true
      $app.ActiveDocument.SaveAs($p, $psdOpts, $true)
      J @{ result='ok'; path=$p }
      break
    }
    'save_as_png' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $p = $env:PSMCP_PATH
      if (-not $p) { $p = Join-Path $env:USERPROFILE "Documents\$($app.ActiveDocument.Name).png" }
      $pngOpts = New-Object -ComObject Photoshop.PNGSaveOptions
      $pngOpts.Compression = 6
      $pngOpts.Interlaced = $false
      $app.ActiveDocument.SaveAs($p, $pngOpts, $true)
      J @{ result='ok'; path=$p }
      break
    }
    'save_as_jpg' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $p = $env:PSMCP_PATH
      $q = [int]$env:PSMCP_QUALITY
      if ($q -lt 1 -or $q -gt 12) { $q = 8 }
      if (-not $p) { $p = Join-Path $env:USERPROFILE "Documents\$($app.ActiveDocument.Name).jpg" }
      $jpgOpts = New-Object -ComObject Photoshop.JPEGSaveOptions
      $jpgOpts.Quality = $q
      $jpgOpts.FormatOptions = 2
      $app.ActiveDocument.SaveAs($p, $jpgOpts, $true)
      J @{ result='ok'; path=$p; quality=$q }
      break
    }
    'resize_document' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $W = [int]$env:PSMCP_W; $H = [int]$env:PSMCP_H
      if ($W -lt 1 -and $H -lt 1) { J @{result='error'; error='invalid size'}; break }
      $d = $app.ActiveDocument
      $w=[double]$d.Width; $h=[double]$d.Height; $r=[double]$d.Resolution
      $pxW = [math]::Round($w*$r); $pxH = [math]::Round($h*$r)
      if ($W -lt 1) { $W = [math]::Round($pxW * ($H / $pxH)) }
      if ($H -lt 1) { $H = [math]::Round($pxH * ($W / $pxW)) }
      $g = Gcd $W $H; if ($g -lt 1) { $g = 72 }
      $d.ResizeImage($W/$g, $H/$g, $g)
      J @{ result='ok'; pixelWidth=$W; pixelHeight=$H }
      break
    }

    # ------ layers ------
    'add_layer' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $nm = $env:PSMCP_NAME
      $d = $app.ActiveDocument
      if ($nm) { $L = $d.ArtLayers.Add($nm) } else { $L = $d.ArtLayers.Add() }
      J @{ result='ok'; layer=$L.Name; totalLayers=$d.Layers.Count }
      break
    }
    'duplicate_layer' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $d = $app.ActiveDocument
      $src = $d.ActiveLayer.Name
      $L = $d.ActiveLayer.Duplicate()
      J @{ result='ok'; from=$src; new=$L.Name; totalLayers=$d.Layers.Count }
      break
    }
    'delete_layer' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $d = $app.ActiveDocument
      $nm = $d.ActiveLayer.Name
      $d.ActiveLayer.Delete()
      J @{ result='ok'; deleted=$nm; totalLayers=$d.Layers.Count }
      break
    }
    'set_layer_opacity' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $o = [int]$env:PSMCP_OPACITY
      if ($o -lt 0) { $o = 0 }; if ($o -gt 100) { $o = 100 }
      $d = $app.ActiveDocument
      $d.ActiveLayer.Opacity = $o
      J @{ result='ok'; layer=$d.ActiveLayer.Name; opacity=$o }
      break
    }

    # ------ text ------
    'add_text_layer' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $txt = $env:PSMCP_TEXT
      $size = [double]$env:PSMCP_SIZE
      if ($size -lt 1) { $size = 48 }
      $x = [double]$env:PSMCP_X
      $y = [double]$env:PSMCP_Y
      $r = [int]$env:PSMCP_R; $g = [int]$env:PSMCP_G; $b = [int]$env:PSMCP_B
      $d = $app.ActiveDocument
      $ppi = [double]$d.Resolution
      $xi = $x / $ppi; $yi = $y / $ppi
      $tf = $d.ArtLayers.Add()
      $tf.Kind = 2  # psTextLayer
      $tf.TextItem.Contents = $txt
      $tf.TextItem.Size = $size
      try { $tf.TextItem.Position = @($xi, $yi) } catch {}
      try {
        if ($r -ge 0 -and $g -ge 0 -and $b -ge 0) {
          $col = NewRgbColor $r $g $b
          $tf.TextItem.Color = $col
        }
      } catch {}
      J @{ result='ok'; layer=$tf.Name; text=$txt; size=$size; x=$x; y=$y }
      break
    }

    # ------ color & fill ------
    'set_foreground_color' {
      $r = [int]$env:PSMCP_R; $g = [int]$env:PSMCP_G; $b = [int]$env:PSMCP_B
      $app.ForegroundColor = (NewRgbColor $r $g $b)
      J @{ result='ok'; r=$r; g=$g; b=$b }
      break
    }
    'fill_layer' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $r = [int]$env:PSMCP_R; $g = [int]$env:PSMCP_G; $b = [int]$env:PSMCP_B
      $d = $app.ActiveDocument
      $app.ForegroundColor = (NewRgbColor $r $g $b)
      $d.Selection.SelectAll()
      $d.Selection.Fill(1)
      $d.Selection.Deselect()
      J @{ result='ok'; r=$r; g=$g; b=$b }
      break
    }

    # ------ filters ------
    'apply_gaussian_blur' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $radius = [double]$env:PSMCP_RADIUS
      if ($radius -lt 0.1) { $radius = 5.0 }
      $app.ActiveDocument.ActiveLayer.ApplyGaussianBlur($radius)
      J @{ result='ok'; radius=$radius }
      break
    }
    'apply_unsharp_mask' {
      if ($app.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $amt = [double]$env:PSMCP_AMOUNT; if ($amt -lt 1) { $amt = 100 }
      $rad = [double]$env:PSMCP_RADIUS; if ($rad -lt 0.1) { $rad = 2.0 }
      $thr = [double]$env:PSMCP_THRESHOLD; if ($thr -lt 0) { $thr = 0 }
      $app.ActiveDocument.ActiveLayer.ApplyUnSharpMask($amt, $rad, $thr)
      J @{ result='ok'; amount=$amt; radius=$rad; threshold=$thr }
      break
    }

    # ------ actions ------
    'do_action' {
      $a = $env:PSMCP_ACTION; $f = $env:PSMCP_FROM
      if (-not $a) { J @{result='error'; error='missing action name'}; break }
      if ($f) { $app.DoAction($a, $f) } else { $app.DoAction($a) }
      J @{ result='ok'; action=$a; from=$f }
      break
    }

    default { J @{result='error'; error="unknown tool: $env:PSMCP_TOOL"} }
  }
}
catch {
  J @{ result='error'; error=$_.Exception.Message }
}
