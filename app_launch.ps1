# SEO Toolkit Pro - open the Edge app window and give it its OWN taskbar identity.
# Edge --app windows normally group under Edge's icon. The reliable fix is to stamp an
# AppUserModelID + RelaunchIconResource directly onto the window after it appears
# (SHGetPropertyStoreForWindow), which makes the taskbar show OUR icon and group it
# separately. Best-effort: any failure is silent and the window still works.
param(
  [Parameter(Mandatory=$true)][string]$Edge,
  [Parameter(Mandatory=$true)][string]$Url,
  [Parameter(Mandatory=$true)][string]$Profile,
  [Parameter(Mandatory=$true)][string]$Icon,
  [Parameter(Mandatory=$true)][string]$WorkDir
)
$ErrorActionPreference = 'Stop'
$aumid = 'DigiGyan.SEOToolkitPro'
$edgeArgs = '--app=' + $Url + ' --user-data-dir="' + $Profile + '" --window-size=1100,820 --no-first-run --no-default-browser-check'

try { Start-Process -FilePath $Edge -ArgumentList $edgeArgs } catch { exit 0 }

try {
  Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
namespace STPwin {
  [StructLayout(LayoutKind.Sequential)] public struct PropertyKey { public Guid fmtid; public uint pid; }
  [StructLayout(LayoutKind.Sequential)] public struct PropVariant { public ushort vt; public ushort r1; public ushort r2; public ushort r3; public IntPtr p; public int p2; }
  [ComImport, Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IPropertyStore {
    void GetCount(out uint c); void GetAt(uint i, out PropertyKey pk);
    void GetValue(ref PropertyKey k, out PropVariant pv);
    void SetValue(ref PropertyKey k, ref PropVariant pv); void Commit();
  }
  public static class Native {
    [DllImport("shell32.dll")] public static extern int SHGetPropertyStoreForWindow(IntPtr hwnd, ref Guid riid, out IPropertyStore pps);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    public delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
    static Guid IID = new Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99");
    static PropertyKey PK(uint pid){ return new PropertyKey{ fmtid = new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), pid = pid }; }
    static void SetStr(IPropertyStore ps, uint pid, string val){
      var pv = new PropVariant{ vt = 31, p = Marshal.StringToCoTaskMemUni(val) };
      var k = PK(pid); ps.SetValue(ref k, ref pv); Marshal.FreeCoTaskMem(pv.p);
    }
    public static bool Apply(string title, string aumid, string iconRes, string relaunchCmd, string displayName){
      IntPtr found = IntPtr.Zero;
      EnumWindows((h,l)=>{
        if(!IsWindowVisible(h)) return true;
        var sb = new StringBuilder(300); GetWindowText(h, sb, 300);
        if(sb.ToString() == title){ found = h; return false; }
        return true;
      }, IntPtr.Zero);
      if(found == IntPtr.Zero) return false;
      IPropertyStore ps;
      if(SHGetPropertyStoreForWindow(found, ref IID, out ps) != 0 || ps == null) return false;
      SetStr(ps, 5, aumid);         // System.AppUserModel.ID
      SetStr(ps, 3, iconRes);       // System.AppUserModel.RelaunchIconResource
      SetStr(ps, 2, relaunchCmd);   // System.AppUserModel.RelaunchCommand
      SetStr(ps, 4, displayName);   // System.AppUserModel.RelaunchDisplayNameResource
      ps.Commit();
      Marshal.ReleaseComObject(ps);
      return true;
    }
  }
}
'@
  $iconRes  = $Icon + ',0'
  $relaunch = 'wscript.exe "' + (Join-Path $WorkDir 'Start Tool.vbs') + '"'
  for($i = 0; $i -lt 40; $i++){
    Start-Sleep -Milliseconds 300
    if([STPwin.Native]::Apply('SEO Toolkit Pro', $aumid, $iconRes, $relaunch, 'SEO Toolkit Pro')){ break }
  }
} catch {}
