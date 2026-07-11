# SEO Toolkit Pro - startup splash. Shows a small "Starting..." window immediately so
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
  $bar.Minimum = 0; $bar.Maximum = 100; $bar.Value = 0
  $bar.Dock = 'Bottom'; $bar.Height = 10
  $f.Controls.Add($bar)

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.Text      = "Starting SEO Toolkit Pro`r`nPlease wait a moment... 0%"
  $lbl.ForeColor = [System.Drawing.Color]::White
  $lbl.Font      = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
  $lbl.TextAlign = 'MiddleCenter'
  $lbl.Dock      = 'Fill'
  $f.Controls.Add($lbl)

  # A real percentage, not just a bouncing bar - approximated from elapsed time toward
  # the same timeout this splash already waits on, since the server doesn't report real
  # milestones back here. Capped at 99% until the port file actually appears so it never
  # falsely claims "100%" before the app is really ready.
  $script:ms = 0
  $t = New-Object System.Windows.Forms.Timer
  $t.Interval = 400
  $t.Add_Tick({
    $script:ms += 400
    $portReady = Test-Path $portFile
    $pct = [Math]::Min(99, [Math]::Floor(($script:ms / 50000.0) * 100))
    if ($portReady) { $pct = 100 }
    $bar.Value = $pct
    $lbl.Text = "Starting SEO Toolkit Pro`r`nPlease wait a moment... $pct%"
    if ($portReady -or ($script:ms -gt 50000)) { $t.Stop(); $f.Close() }
  })
  $t.Start()

  [System.Windows.Forms.Application]::Run($f)
} catch {}
