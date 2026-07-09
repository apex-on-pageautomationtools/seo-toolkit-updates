# SEO Toolkit Pro — startup splash. Shows a small "Starting…" window immediately so
# the user has feedback while the server comes up, and closes itself the moment the
# server writes its port file (or after a timeout). Purely cosmetic — any failure is
# silent and does not affect launching.
try {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  $portFile = Join-Path $env:TEMP ".grc_port"

  $f = New-Object System.Windows.Forms.Form
  $f.FormBorderStyle = 'None'
  $f.StartPosition   = 'CenterScreen'
  $f.Size            = New-Object System.Drawing.Size(370, 140)
  $f.BackColor       = [System.Drawing.Color]::FromArgb(26, 42, 65)
  $f.TopMost         = $true
  $f.ShowInTaskbar   = $false

  $bar = New-Object System.Windows.Forms.ProgressBar
  $bar.Style = 'Marquee'; $bar.MarqueeAnimationSpeed = 30
  $bar.Dock = 'Bottom'; $bar.Height = 10
  $f.Controls.Add($bar)

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.Text      = "Starting SEO Toolkit Pro`r`nPlease wait a moment…"
  $lbl.ForeColor = [System.Drawing.Color]::White
  $lbl.Font      = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
  $lbl.TextAlign = 'MiddleCenter'
  $lbl.Dock      = 'Fill'
  $f.Controls.Add($lbl)

  $script:ms = 0
  $t = New-Object System.Windows.Forms.Timer
  $t.Interval = 400
  $t.Add_Tick({
    $script:ms += 400
    if ((Test-Path $portFile) -or ($script:ms -gt 50000)) { $t.Stop(); $f.Close() }
  })
  $t.Start()

  [System.Windows.Forms.Application]::Run($f)
} catch {}
