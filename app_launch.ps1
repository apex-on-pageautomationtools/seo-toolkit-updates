# SEO Toolkit Pro - launch the Edge app window with its OWN taskbar identity.
# Edge's --app windows normally group under Edge's icon. Launching through a .lnk
# that carries a custom AppUserModelID (AUMID) + icon makes Windows show OUR icon
# and group it separately -- the same mechanism Chrome/Edge PWA shortcuts use.
param(
  [Parameter(Mandatory=$true)][string]$Edge,
  [Parameter(Mandatory=$true)][string]$Url,
  [Parameter(Mandatory=$true)][string]$Profile,
  [Parameter(Mandatory=$true)][string]$Icon,
  [Parameter(Mandatory=$true)][string]$WorkDir
)

$ErrorActionPreference = 'Stop'
$aumid   = 'DigiGyan.SEOToolkitPro'
$edgeArgs = '--app=' + $Url + ' --user-data-dir="' + $Profile + '" --window-size=1100,820 --no-first-run --no-default-browser-check'

function New-AumidShortcut {
  param($LnkPath, $Target, $Arguments, $IconPath, $WorkingDir, $AppId)

  # 1) Base shortcut (target / args / icon) via WScript.Shell
  $sh = New-Object -ComObject WScript.Shell
  $s  = $sh.CreateShortcut($LnkPath)
  $s.TargetPath       = $Target
  $s.Arguments        = $Arguments
  $s.IconLocation     = "$IconPath,0"
  $s.WorkingDirectory = $WorkingDir
  $s.Description      = 'SEO Toolkit Pro'
  $s.Save()

  # 2) Stamp the AppUserModelID onto the shortcut's property store
  Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace STP {
  [ComImport, Guid("00021401-0000-0000-C000-000000000046")] public class CShellLink {}

  [ComImport, Guid("0000010b-0000-0000-C000-000000000046"),
   InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IPersistFile {
    void GetClassID(out Guid pClassID);
    [PreserveSig] int IsDirty();
    void Load([MarshalAs(UnmanagedType.LPWStr)] string f, int m);
    void Save([MarshalAs(UnmanagedType.LPWStr)] string f, [MarshalAs(UnmanagedType.Bool)] bool remember);
    void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string f);
    void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string f);
  }

  [StructLayout(LayoutKind.Sequential)]
  public struct PropertyKey { public Guid fmtid; public uint pid; }

  [StructLayout(LayoutKind.Sequential)]
  public struct PropVariant {
    public ushort vt; public ushort r1; public ushort r2; public ushort r3;
    public IntPtr p; public int p2;
  }

  [ComImport, Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"),
   InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IPropertyStore {
    void GetCount(out uint c);
    void GetAt(uint i, out PropertyKey pk);
    void GetValue(ref PropertyKey k, out PropVariant pv);
    void SetValue(ref PropertyKey k, ref PropVariant pv);
    void Commit();
  }

  public static class Lnk {
    static PropertyKey PKEY_AppUserModel_ID = new PropertyKey {
      fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), pid = 5 };

    public static void SetAumid(string lnkPath, string appId) {
      var link  = new CShellLink();
      var pf     = (IPersistFile)link;
      pf.Load(lnkPath, 2 /*STGM_READWRITE - required so Save() can persist*/);
      var store = (IPropertyStore)link;
      var pv = new PropVariant { vt = 31 /*VT_LPWSTR*/, p = Marshal.StringToCoTaskMemUni(appId) };
      var key = PKEY_AppUserModel_ID;
      store.SetValue(ref key, ref pv);
      store.Commit();
      Marshal.FreeCoTaskMem(pv.p);
      pf.Save(lnkPath, true);
      Marshal.ReleaseComObject(link);
    }
  }
}
'@
  [STP.Lnk]::SetAumid($LnkPath, $AppId)
}

try {
  $lnkDir = Join-Path $env:LOCALAPPDATA 'SEO Toolkit Pro'
  if (-not (Test-Path $lnkDir)) { New-Item -ItemType Directory -Path $lnkDir -Force | Out-Null }
  $lnk = Join-Path $lnkDir 'SEO Toolkit Pro.lnk'

  New-AumidShortcut $lnk $Edge $edgeArgs $Icon $WorkDir $aumid

  # Launch via the shortcut so the Edge window inherits the AUMID + icon
  Start-Process -FilePath $lnk
  exit 0
}
catch {
  # Guaranteed fallback: launch Edge directly (window works, taskbar shows Edge icon)
  Start-Process -FilePath $Edge -ArgumentList $edgeArgs
  exit 0
}
