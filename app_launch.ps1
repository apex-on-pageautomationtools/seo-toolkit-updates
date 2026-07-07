# SEO Toolkit Pro - simple fallback that opens the app window in Edge. Kept minimal
# and benign (no window/Win32 manipulation) so antivirus tools don't flag it — an
# earlier version used low-level window APIs to set the taskbar icon and got
# quarantined by Quick Heal. The launcher now opens Edge directly; this remains only
# as a harmless fallback.
param(
  [string]$Edge,
  [string]$Url,
  [string]$Profile,
  [string]$Icon,
  [string]$WorkDir
)
try {
  $a = '--app=' + $Url + ' --user-data-dir="' + $Profile + '" --window-size=1100,820 --no-first-run --no-default-browser-check'
  Start-Process -FilePath $Edge -ArgumentList $a
} catch {}
