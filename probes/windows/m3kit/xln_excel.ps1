
# xln: drive desktop Excel through COM. Checked on Windows 11 (M3 check, 2026-10-05).
param([string]$Action, [string]$Path)
$ErrorActionPreference = 'Stop'
$full = [System.IO.Path]::GetFullPath($Path)
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $xl = $null }
# A workbook opened from Explorer can live in another Excel process than the one
# GetActiveObject returns (measured: R08). Excel registers each open workbook in the
# Running Object Table under its path, so the table finds it in any Excel instance.
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class XlnRot {
  [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable rot);
  [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved, out IBindCtx ctx);
  public static object Find(string path) {
    IRunningObjectTable rot; IBindCtx ctx; IEnumMoniker en;
    if (GetRunningObjectTable(0, out rot) != 0 || CreateBindCtx(0, out ctx) != 0) return null;
    rot.EnumRunning(out en);
    IMoniker[] m = new IMoniker[1];
    while (en.Next(1, m, IntPtr.Zero) == 0) {
      string name;
      try { m[0].GetDisplayName(ctx, null, out name); } catch { continue; }
      if (string.Equals(name, path, StringComparison.OrdinalIgnoreCase)) {
        object o;
        if (rot.GetObject(m[0], out o) == 0) return o;
      }
    }
    return null;
  }
  [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
  [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  // Restores a minimised window and asks for the foreground; returns the window's process id.
  public static uint Front(long hwnd) {
    IntPtr h = new IntPtr(hwnd);
    if (IsIconic(h)) ShowWindow(h, 9);
    SetForegroundWindow(h);
    uint pid; GetWindowThreadProcessId(h, out pid); return pid;
  }
}
"@
# Excel's window to the front after an open, so the author sees the workbook. Windows may
# only flash the taskbar button (foreground lock): that is fine. Never fails the open.
function Show-Book($wb) {
  try {
    $app = $wb.Application
    $app.Visible = $true
    try { $wb.Activate() } catch {}
    $xpid = [XlnRot]::Front([long]$app.Hwnd)
    $null = (New-Object -ComObject WScript.Shell).AppActivate([int]$xpid)
  } catch {}
}
function Find-Book($xl, $full) {
  if ($xl -ne $null) { foreach ($wb in $xl.Workbooks) { if ($wb.FullName -ieq $full) { return $wb } } }
  $o = [XlnRot]::Find($full)
  if ($o -ne $null) { try { if ($o.FullName -ieq $full) { return $o } } catch {} }
  return $null
}
switch ($Action) {
  'state' {
    $wb = Find-Book $xl $full
    if ($wb -eq $null -and $xl -eq $null) { 'notrunning'; break }
    if ($wb -eq $null) { 'closed' } elseif ($wb.Saved) { 'saved' } else { 'unsaved' }
  }
  'close' {
    $wb = Find-Book $xl $full
    if ($wb -eq $null) { 'not open' } else { $wb.Close($false); 'closed' }
  }
  'open' {
    $wb = Find-Book $xl $full
    if ($wb -ne $null) { Show-Book $wb; 'opened'; break }
    if ($xl -eq $null) { $xl = New-Object -ComObject Excel.Application }
    # Visible and under the user's control, so Excel stays open when this script ends.
    $xl.Visible = $true
    $xl.UserControl = $true
    # The plain call: measured on Windows (M3 check, 2026-10-05), Excel answers a file that
    # needs repair with an error here, no prompt, while the 15-argument form with
    # [Type]::Missing and CorruptLoad failed on every file, clean ones too.
    $wb = $xl.Workbooks.Open($full)
    if ($wb -ne $null -and $wb.FullName -ieq $full) { Show-Book $wb; 'opened' } else { 'missing' }
  }
}
