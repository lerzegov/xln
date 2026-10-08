// E7, reopening the workbook in desktop Excel around a build: close it (without saving),
// let the caller write the file, open it again. On the Mac through AppleScript, on Windows
// through Excel's COM interface from PowerShell. This module holds the scripts and what
// their answers mean; running a program is the caller's (the CLI's and the desktop
// extension's: the core stays free of Node), passed in as `exec`.
//
// The Windows script has NOT been run yet (written on a Mac): treat it as a draft until
// someone checks it on a lab PC (PLAN.md lists the check).
//
// Excel may answer an open with its repair prompt (probe F8: a stale calcChain). The prompt
// blocks Excel, and System Events cannot press its buttons without Accessibility rights,
// so an open is given a time limit and then checked: a workbook of the expected name must
// be open. Nothing is retried: a second open would stack a second prompt.

export interface ExecResult {
  /** Standard output, trimmed. */
  out: string;
  timedOut: boolean;
  /** Set when the program failed to start or exited with an error. */
  error?: string;
}

/** Runs `osascript -` with `script` on standard input, or PowerShell with `script` as a `.ps1` file. */
export type ExcelExec = (program: "osascript" | "powershell", script: string, args: string[], timeoutMs: number) => ExecResult;

export interface ExcelWorkbookState {
  /** Excel answered (it is running). */
  running: boolean;
  /** Excel has this file open. */
  open: boolean;
  /** The open workbook has no unsaved changes (undefined when it is not open). */
  saved?: boolean;
}

export interface ExcelResult {
  ok: boolean;
  message: string;
}

export interface ExcelControl {
  state(path: string): ExcelWorkbookState;
  /** Closes the workbook without saving. */
  close(path: string): ExcelResult;
  /** Opens the workbook, checks it opened (no repair prompt in the way), and brings Excel to the front. */
  open(path: string): ExcelResult;
}

// macOS. Workbooks are found by file name (Excel cannot open two of the same name) and
// then compared by full name, which Excel for Mac reports as a POSIX path. `state` first
// asks whether Excel runs (`running` needs no System Events), so it never launches Excel.
export const MAC_STATE_SCRIPT = `
on run argv
  set p to item 1 of argv
  set n to item 2 of argv
  if not (application "Microsoft Excel" is running) then return "notrunning"
  tell application "Microsoft Excel"
    repeat with i from 1 to (count of workbooks)
      if (name of workbook i) is n then
        if (full name of workbook i) is not p then return "other"
        if saved of workbook i then return "saved"
        return "unsaved"
      end if
    end repeat
  end tell
  return "closed"
end run`;

export const MAC_CLOSE_SCRIPT = `
on run argv
  set p to item 1 of argv
  set n to item 2 of argv
  tell application "Microsoft Excel"
    repeat with i from 1 to (count of workbooks)
      if (name of workbook i) is n and (full name of workbook i) is p then
        close workbook i saving no
        return "closed"
      end if
    end repeat
  end tell
  return "not open"
end run`;

export const MAC_OPEN_SCRIPT = `
on run argv
  set p to item 1 of argv
  set n to item 2 of argv
  tell application "Microsoft Excel"
    with timeout of 60 seconds
      open workbook workbook file name p
    end timeout
  end tell
  -- Checked again after a pause, with a short time limit: a dialog still up (a repair
  -- prompt, a file-access request) blocks Excel, and the second look times out.
  delay 2
  tell application "Microsoft Excel"
    with timeout of 15 seconds
      repeat with i from 1 to (count of workbooks)
        if (name of workbook i) is n and (full name of workbook i) is p then
          -- To the front, so the author sees the rebuilt workbook (after the check: an
          -- activation never stands in for it).
          activate
          return "opened"
        end if
      end repeat
    end timeout
  end tell
  return "missing"
end run`;

export const WINDOWS_EXCEL_SCRIPT = String.raw`
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
`;

function baseName(path: string): string {
  const i = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  return i < 0 ? path : path.slice(i + 1);
}

const BUSY = "Excel did not answer within 30 s (a dialog may be open in Excel)";
const SLOW_OPEN = "Excel did not finish opening within 90 s: look for a dialog in Excel (a repair prompt: answer No, then restore the backup)";

/**
 * The control for `platform` (Node's platform name), or undefined where there is no desktop
 * Excel. `repairLogsSince` (macOS) lists Excel's repair logs for a file written since a time.
 */
export function excelControl(platform: string, exec: ExcelExec, repairLogsSince?: (path: string, sinceMs: number) => string[]): ExcelControl | undefined {
  let run: (action: "state" | "close" | "open", path: string, timeoutMs: number) => ExecResult;
  if (platform === "darwin") {
    const scripts = { state: MAC_STATE_SCRIPT, close: MAC_CLOSE_SCRIPT, open: MAC_OPEN_SCRIPT };
    run = (action, path, t) => exec("osascript", scripts[action], [path, baseName(path)], t);
  } else if (platform === "win32") {
    run = (action, path, t) => exec("powershell", WINDOWS_EXCEL_SCRIPT, ["-Action", action, "-Path", path], t);
  } else return undefined;

  return {
    state(path) {
      const r = run("state", path, 20_000);
      if (r.out === "notrunning") return { running: false, open: false };
      if (r.timedOut || r.error) return { running: true, open: false };
      if (r.out === "saved" || r.out === "unsaved") return { running: true, open: true, saved: r.out === "saved" };
      return { running: true, open: false };
    },
    close(path) {
      const r = run("close", path, 30_000);
      if (r.timedOut) return { ok: false, message: BUSY };
      if (r.error) return { ok: false, message: `Excel could not close the workbook: ${r.error}` };
      return { ok: true, message: r.out === "closed" ? "closed in Excel without saving" : "was not open in Excel" };
    },
    open(path) {
      const t0 = Date.now() - 1000;
      const r = run("open", path, 90_000);
      const logs = repairLogsSince?.(path, t0) ?? [];
      if (logs.length) return { ok: false, message: `Excel repaired the file on opening (log: ${logs[0]}): restore the backup and report this` };
      if (r.timedOut) return { ok: false, message: SLOW_OPEN };
      if (r.error && /-1712|timed out|scaduto/i.test(r.error)) return { ok: false, message: "Excel did not answer after opening the workbook: a dialog is waiting in Excel (a repair prompt: answer No and restore the backup; a file-access request: grant it)" };
      if (r.error) return { ok: false, message: `Excel could not open the workbook (if it needs repair, restore the backup): ${r.error}` };
      if (r.out !== "opened") return { ok: false, message: "Excel did not open the workbook (a repair prompt or a file-access request may be waiting in Excel)" };
      return { ok: true, message: "opened in Excel" };
    },
  };
}
