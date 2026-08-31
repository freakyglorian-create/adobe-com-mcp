# Illustrator COM automation dispatcher.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function J($o) { $o | ConvertTo-Json -Compress -Depth 4 }

function NewRgbColor($r, $g, $b) {
  $c = New-Object -ComObject Illustrator.RGBColor
  $c.Red = [double]$r
  $c.Green = [double]$g
  $c.Blue = [double]$b
  return $c
}

try {
  $ai = New-Object -ComObject Illustrator.Application

  switch ($env:AIMCP_TOOL) {

    # ------ document ------
    'ai_create_document' {
      $W = [double]$env:AIMCP_W; $H = [double]$env:AIMCP_H
      if ($W -le 0) { $W = 612 }
      if ($H -le 0) { $H = 792 }
      $doc = $ai.Documents.Add($W, $H)
      J @{result='ok'; name=$doc.Name; width=$W; height=$H; unit='points'}
      break
    }
    'ai_get_active_info' {
      if ($ai.Documents.Count -eq 0) { J @{result='ok'; active=$false}; break }
      $d = $ai.ActiveDocument
      J @{result='ok'; name=$d.Name; width=$d.Width; height=$d.Height
           paths=$d.PathItems.Count; textFrames=$d.TextFrames.Count; open=$ai.Documents.Count}
      break
    }
    'ai_list_documents' {
      $list = @()
      for ($i = 1; $i -le $ai.Documents.Count; $i++) { $list += @{name=$ai.Documents.Item($i).Name} }
      J @{result='ok'; count=$ai.Documents.Count; documents=$list}
      break
    }
    'ai_close_document' {
      if ($ai.Documents.Count -eq 0) { J @{result='ok'; closed=$false}; break }
      $nm = $ai.ActiveDocument.Name
      $ai.ActiveDocument.Close(2)  # aiDoNotSaveChanges
      J @{result='ok'; closed=$true; name=$nm}
      break
    }

    # ------ shapes ------
    'ai_add_rectangle' {
      if ($ai.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $x = [double]$env:AIMCP_X; $y = [double]$env:AIMCP_Y
      $w = [double]$env:AIMCP_W; $h = [double]$env:AIMCP_H
      if ($w -le 0) { $w = 200 }; if ($h -le 0) { $h = 150 }
      $d = $ai.ActiveDocument
      $r = $d.PathItems.Rectangle($x, $y, $w, $h)
      # fill color if given
      $fr = [int]$env:AIMCP_FR; $fg = [int]$env:AIMCP_FG; $fb = [int]$env:AIMCP_FB
      if ($fr -ge 0 -and $fg -ge 0 -and $fb -ge 0) {
        $r.Filled = $true
        $r.FillColor = (NewRgbColor $fr $fg $fb)
      }
      J @{result='ok'; name=$r.Name; x=$x; y=$y; width=$w; height=$h}
      break
    }
    'ai_add_ellipse' {
      if ($ai.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $x = [double]$env:AIMCP_X; $y = [double]$env:AIMCP_Y
      $w = [double]$env:AIMCP_W; $h = [double]$env:AIMCP_H
      if ($w -le 0) { $w = 150 }; if ($h -le 0) { $h = 100 }
      $d = $ai.ActiveDocument
      $e = $d.PathItems.Ellipse($x, $y, $w, $h)
      $fr = [int]$env:AIMCP_FR; $fg = [int]$env:AIMCP_FG; $fb = [int]$env:AIMCP_FB
      if ($fr -ge 0 -and $fg -ge 0 -and $fb -ge 0) {
        $e.Filled = $true
        $e.FillColor = (NewRgbColor $fr $fg $fb)
      }
      J @{result='ok'; name=$e.Name; x=$x; y=$y; width=$w; height=$h}
      break
    }
    'ai_add_polygon' {
      if ($ai.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $x = [double]$env:AIMCP_X; $y = [double]$env:AIMCP_Y
      $r = [double]$env:AIMCP_RADIUS
      $s = [int]$env:AIMCP_SIDES
      if ($r -le 0) { $r = 100 }; if ($s -lt 3) { $s = 6 }
      $d = $ai.ActiveDocument
      $p = $d.PathItems.Polygon($x, $y, $r, $s)
      $fr = [int]$env:AIMCP_FR; $fg = [int]$env:AIMCP_FG; $fb = [int]$env:AIMCP_FB
      if ($fr -ge 0 -and $fg -ge 0 -and $fb -ge 0) {
        $p.Filled = $true
        $p.FillColor = (NewRgbColor $fr $fg $fb)
      }
      J @{result='ok'; name=$p.Name; x=$x; y=$y; radius=$r; sides=$s}
      break
    }

    # ------ text ------
    'ai_add_text' {
      if ($ai.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $txt = $env:AIMCP_TEXT
      $x = [double]$env:AIMCP_X; $y = [double]$env:AIMCP_Y
      $size = [double]$env:AIMCP_SIZE
      if (-not $txt) { $txt = "Text" }
      if ($size -lt 1) { $size = 24 }
      $d = $ai.ActiveDocument
      $tf = $d.TextFrames.Add()
      $tf.Contents = $txt
      try { $tf.TextRange.CharacterAttributes.Size = $size } catch {}
      try {
        $fr = [int]$env:AIMCP_FR; $fg = [int]$env:AIMCP_FG; $fb = [int]$env:AIMCP_FB
        if ($fr -ge 0 -and $fg -ge 0 -and $fb -ge 0) {
          $tf.TextRange.CharacterAttributes.FillColor = (NewRgbColor $fr $fg $fb)
        }
      } catch {}
      try {
        $tf.Top = $y
        $tf.Left = $x
      } catch {}
      J @{result='ok'; text=$txt; size=$size; x=$x; y=$y}
      break
    }

    # ------ save/export ------
    'ai_save_as_ai' {
      if ($ai.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $p = $env:AIMCP_PATH
      if (-not $p) { $p = Join-Path $env:USERPROFILE "Documents\$($ai.ActiveDocument.Name).ai" }
      $ai.ActiveDocument.SaveAs($p)
      J @{result='ok'; path=$p}
      break
    }
    'ai_export_svg' {
      if ($ai.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $p = $env:AIMCP_PATH
      if (-not $p) { $p = Join-Path $env:USERPROFILE "Documents\$($ai.ActiveDocument.Name).svg" }
      $opts = New-Object -ComObject Illustrator.SVGExportOptions
      $opts.FontType = 1  # aiSVGFontOutline
      $ai.ActiveDocument.Export($p, 10, $opts)  # 10 = aiSVG
      J @{result='ok'; path=$p}
      break
    }
    'ai_export_png' {
      if ($ai.Documents.Count -eq 0) { J @{result='error'; error='no document'}; break }
      $p = $env:AIMCP_PATH
      if (-not $p) { $p = Join-Path $env:USERPROFILE "Documents\$($ai.ActiveDocument.Name).png" }
      $opts = New-Object -ComObject Illustrator.PNGExportOptions
      $opts.AntiAliasing = $true
      $opts.Transparency = $true
      $ai.ActiveDocument.Export($p, 4, $opts)  # 4 = aiPNG
      J @{result='ok'; path=$p}
      break
    }

    default { J @{result='error'; error="unknown tool: $env:AIMCP_TOOL"} }
  }
}
catch {
  J @{result='error'; error=$_.Exception.Message}
}
