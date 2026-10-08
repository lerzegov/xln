# m3_check.ps1 -- M3 Windows check: open the kit workbooks that xln wrote (M3a names,
# M3b cell formulas, D5 embedded source, D6 tags, F8 files) in Windows Excel, compare what
# Excel shows with m3kit\*.expected.json, save a copy of each, and try the PowerShell
# script behind `xln build --reopen` (E7).
#
#   run_m3_check.cmd                     (double-click; uses -ExecutionPolicy Bypass)
#   powershell -NoProfile -ExecutionPolicy Bypass -File m3_check.ps1 [-SkipReopen | -ReopenOnly]
#
# Part 1 starts its OWN Excel instance (your other workbooks are not touched), opens each
# kit workbook, reads cells (Formula2, Value2, Text), names (RefersTo, Comment, Visible),
# expressions (Evaluate) and the embedded-source part (CustomXMLParts), saves a copy as
# probes\results\m3win\<name>_winsaved.xlsx and closes it. The two files that should make
# Excel ask about repairing come last, each in a fresh instance, after a message saying
# what to click. Part 2 needs Excel closed: it runs m3kit\xln_excel.ps1 (the exact
# script xln writes) the way xln does, on copies in %LOCALAPPDATA%\xln-m3check.
#
# Output: "ID | STATUS | detail" lines, written to probes\results\m3-win-<host>-<time>.txt.
# This file is ASCII only: Windows PowerShell 5.1 reads a script without a BOM as ANSI.

param([switch]$SkipReopen, [switch]$ReopenOnly)

$ErrorActionPreference = 'Stop'
$script:rpt = New-Object System.Collections.Generic.List[string]
$script:count = @{ PASS = 0; FAIL = 0; WARN = 0 }
$script:fails = New-Object System.Collections.Generic.List[string]
function Add-Line($tid, $verdict, $det) {
    $line = "$tid | $verdict | " + ("$det" -replace '\s*[\r\n]+\s*', ' / ')
    $script:rpt.Add($line)
    if ($script:count.ContainsKey($verdict)) { $script:count[$verdict]++ }
    if ($verdict -eq 'FAIL' -or $verdict -eq 'WARN') { $script:fails.Add($line) }
    $color = switch ($verdict) { 'PASS' { 'Green' } 'FAIL' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
}
function Check($tid, $ok, $det) { Add-Line $tid ($(if ($ok) { 'PASS' } else { 'FAIL' })) $det }
function Say($text) { Write-Host ''; Write-Host $text -ForegroundColor Cyan }
function Short($s, $n = 120) { if ($null -eq $s) { return '<null>' }; $t = "$s"; if ($t.Length -gt $n) { $t.Substring(0, $n) + '...' } else { $t } }
function Msg($e) { Short ($e.Exception.GetBaseException().Message) 200 }

$here = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$kitDir = Join-Path $here 'm3kit'
$resDir = [System.IO.Path]::GetFullPath((Join-Path $here '..\results'))
$winDir = Join-Path $resDir 'm3win'
if (-not (Test-Path $winDir)) { New-Item -ItemType Directory $winDir | Out-Null }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$outFile = Join-Path $resDir ("m3-win-{0}-{1}.txt" -f $env:COMPUTERNAME, $stamp)
function Save-Report { $script:rpt | Set-Content -Encoding UTF8 $outFile }
$inv = [System.Globalization.CultureInfo]::InvariantCulture
# Excel opens copies, never the kit files themselves: a repo inside OneDrive could
# AutoSave them. Copies are unblocked in case the repo came as a downloaded zip.
$stage = Join-Path $env:LOCALAPPDATA 'xln-m3check'
function Stage-Copy($src, $leaf) {
    if (-not (Test-Path $stage)) { New-Item -ItemType Directory $stage | Out-Null }
    $dst = Join-Path $stage $leaf
    Copy-Item $src $dst -Force
    try { Unblock-File $dst } catch {}
    return $dst
}

if (-not (Test-Path (Join-Path $kitDir 'kit.json'))) {
    Write-Host "m3kit\kit.json not found next to this script ($kitDir). Run this from the repo's probes\windows folder." -ForegroundColor Red
    return
}
$kit = Get-Content -Raw -Encoding UTF8 (Join-Path $kitDir 'kit.json') | ConvertFrom-Json

Add-Line 'H00' 'INFO' ("# xln M3 Windows check, script v1  {0:u}" -f (Get-Date).ToUniversalTime())
Add-Line 'E01' 'INFO' ("OS {0}; PowerShell {1}; culture {2}" -f [Environment]::OSVersion.VersionString, $PSVersionTable.PSVersion, (Get-Culture).Name)
$lm = $ExecutionContext.SessionState.LanguageMode
Check 'E02' ($lm -eq 'FullLanguage') "LanguageMode=$lm"
Add-Line 'E03' 'INFO' ("kit: {0}; results: {1}" -f $kitDir, $resDir)
if ($env:OneDrive -and $kitDir.StartsWith($env:OneDrive, [System.StringComparison]::OrdinalIgnoreCase)) { Add-Line 'E03' 'INFO' 'the repo is inside the OneDrive folder' }

# A watchdog that prints to this window while a COM call is blocked by a dialog in Excel
# (PowerShell itself cannot print until the call returns). Compiled C#; optional.
$script:dog = $false
try {
    if (-not ('XlnWatchdog' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Threading;
public static class XlnWatchdog {
    static Timer timer;
    static string text;
    public static void Start(int seconds, string message) {
        Stop();
        text = message;
        timer = new Timer(Tick, null, seconds * 1000, 40000);
    }
    static void Tick(object state) {
        try { Console.Beep(880, 250); } catch { }
        Console.WriteLine();
        Console.WriteLine(text);
    }
    public static void Stop() { if (timer != null) { timer.Dispose(); timer = null; } }
}
'@
    }
    $script:dog = $true
} catch { Add-Line 'E04' 'INFO' ("no watchdog (Add-Type failed): " + (Msg $_)) }
function Start-Dog($seconds, $text) { if ($script:dog) { [XlnWatchdog]::Start($seconds, $text) } }
function Stop-Dog { if ($script:dog) { [XlnWatchdog]::Stop() } }
$dialogText = @"
>>> Excel has not answered for a while: it is probably showing a dialog.
>>> Switch to Excel (Alt+Tab) and answer it:
>>>   - 'We found a problem with some content ... Do you want us to try to recover?'
>>>     ['Abbiamo riscontrato un problema ...']: click Yes [Si], then Close [Chiudi].
>>>   - an error message: click OK.
>>> This script continues by itself afterwards and asks you what you saw.
"@

function Excel-Pids { @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }
$startPids = Excel-Pids
$doReopen = -not $SkipReopen
if ($doReopen -and $startPids.Count -gt 0) {
    Say "Part 2 of this check (xln build --reopen) needs Excel to be closed."
    Write-Host "Save your work and close every Excel window now, then press Enter here."
    Write-Host "(Or type S and Enter to skip part 2; part 1 runs either way in its own Excel.)"
    $a = Read-Host
    if ($a -match '^\s*[sS]') { $doReopen = $false; Add-Line 'E05' 'INFO' 'part 2 (--reopen) skipped by the author' }
}

# ---- comparisons ------------------------------------------------------------------------
$errNames = @{ 2000 = '#NULL!'; 2007 = '#DIV/0!'; 2015 = '#VALUE!'; 2023 = '#REF!'; 2029 = '#NAME?'; 2036 = '#NUM!'; 2042 = '#N/A'; 2043 = '#GETTING_DATA'; 2045 = '#SPILL!'; 2046 = '#CONNECT!'; 2047 = '#BLOCKED!'; 2048 = '#UNKNOWN!'; 2049 = '#FIELD!'; 2050 = '#CALC!' }
# Value2 gives numbers as Double, errors as Int32 (CVErr), text as String, empty as null.
function Show-Value($v) {
    if ($null -eq $v) { return '(empty)' }
    if ($v -is [int]) { $n = $v + 2146828288; if ($errNames.ContainsKey($n)) { return $errNames[$n] } else { return "#ERR$n" } }
    if ($v -is [double]) { return $v.ToString('R', $inv) }
    if ($v -is [bool]) { return $(if ($v) { 'TRUE' } else { 'FALSE' }) }
    return '"' + $v + '"'
}
function Show-Expected($e) {
    if ($null -eq $e) { return '(empty)' }
    if ($e -is [string]) { return '"' + $e + '"' }
    if ($e.PSObject.Properties.Name -contains 'error') { return $e.error }
    return ([double]$e).ToString('R', $inv)
}
function Same-Value($e, $v) {
    if ($null -eq $e) { return ($null -eq $v -or ($v -is [string] -and $v -eq '')) }
    if ($e -is [string]) { return ($v -is [string] -and $v -ceq $e) }
    if ($e.PSObject.Properties.Name -contains 'error') { return ((Show-Value $v) -eq $e.error) }
    if (-not ($v -is [double])) { return $false }
    $x = [double]$e
    return ([Math]::Abs($v - $x) -le 1e-9 * [Math]::Max(1.0, [Math]::Abs($x)))
}
# Formulas modulo whitespace and quotes around simple sheet names, any case.
function Squash($f) { (("$f" -replace '\s+', '') -replace "'([A-Za-z_][A-Za-z0-9_.]*)'!", '$1!').ToLowerInvariant() }
function Name-Key($n) { (("$n" -replace "^'([^']+)'!", '$1!')).ToLowerInvariant() }
function Has($o, $p) { $o.PSObject.Properties.Name -contains $p }
function NL($s) { if ($null -eq $s) { $null } else { "$s".Replace("`r`n", "`n").Replace("`r", "`n") } }

# ---- one kit workbook ---------------------------------------------------------------------
function Read-Kit($xl, $wb, $exp, $id) {
    try { Check "$id.S" ($wb.Worksheets.Count -eq $exp.sheets) ("sheets {0} (expected {1})" -f $wb.Worksheets.Count, $exp.sheets) } catch { Check "$id.S" $false (Msg $_) }
    foreach ($c in $exp.cells) {
        $ref = "{0}!{1}" -f $c.sheet, $c.cell
        try {
            $rg = $wb.Worksheets.Item($c.sheet).Range($c.cell)
            $f = $rg.Formula2; $v = $rg.Value2; $t = $rg.Text
            if (Has $c 'formula') { Check "$id.C" ((Squash $f) -eq (Squash $c.formula)) ("{0} Formula2 '{1}' (expected '{2}')" -f $ref, $f, $c.formula) }
            if (Has $c 'value') { Check "$id.V" (Same-Value $c.value $v) ("{0} value {1} text '{2}' (expected {3})" -f $ref, (Show-Value $v), $t, (Show-Expected $c.value)) }
        } catch { Check "$id.C" $false ("{0}: {1}" -f $ref, (Msg $_)) }
    }
    if ($exp.names.Count -gt 0) {
        $names = @{}
        try { foreach ($n in $wb.Names) { $names[(Name-Key $n.Name)] = $n } } catch { Check "$id.N" $false ("cannot list names: " + (Msg $_)) }
        foreach ($e in $exp.names) {
            $n = $names[(Name-Key $e.name)]
            if ((Has $e 'absent') -and $e.absent) { Check "$id.N" ($null -eq $n) ("name {0} absent{1}" -f $e.name, $(if ($n) { ': found ' + $n.RefersTo } else { '' })); continue }
            if ($null -eq $n) { Check "$id.N" $false ("name {0} not found" -f $e.name); continue }
            try {
                $r = $n.RefersTo
                Check "$id.N" ((Squash $r) -eq (Squash $e.refersTo)) ("name {0} RefersTo '{1}' (expected '{2}')" -f $e.name, $r, $e.refersTo)
                if (Has $e 'visible') { Check "$id.N" ([bool]$n.Visible -eq [bool]$e.visible) ("name {0} Visible {1} (expected {2})" -f $e.name, $n.Visible, $e.visible) }
                $cm = $n.Comment
                if (Has $e 'comment') {
                    $ok = (NL $cm) -ceq (NL $e.comment)
                    $how = if ("$cm".Contains("`r`n")) { 'CR LF' } elseif ("$cm".Contains("`n")) { 'LF' } else { 'no line break' }
                    Check "$id.M" $ok ("name {0} Comment ({1} chars, {2}) '{3}'{4}" -f $e.name, "$cm".Length, $how, (Short $cm 300), $(if ($ok) { '' } else { " (expected '" + (Short $e.comment 300) + "')" }))
                } elseif ($cm) { Add-Line "$id.M" 'INFO' ("name {0} Comment ({1} chars) '{2}'" -f $e.name, "$cm".Length, (Short $cm 80)) }
            } catch { Check "$id.N" $false ("name {0}: {1}" -f $e.name, (Msg $_)) }
        }
    }
    foreach ($e in $exp.evaluate) {
        try { $v = $xl.Evaluate($e.expr); Check "$id.E" (Same-Value $e.value $v) ("Evaluate {0} -> {1} (expected {2})" -f $e.expr, (Show-Value $v), (Show-Expected $e.value)) }
        catch { Check "$id.E" $false ("Evaluate {0}: {1}" -f $e.expr, (Msg $_)) }
    }
    try {
        $parts = $wb.CustomXMLParts.SelectByNamespace($kit.embedNamespace)
        $k = $parts.Count
        $det = "custom XML parts with {0}: {1} (expected {2})" -f $kit.embedNamespace, $k, $(if ($exp.customXml) { 1 } else { 0 })
        if ($k -gt 0) { $x = $parts.Item(1).XML; $det += ("; as Excel holds it: {0} chars, {1}" -f $x.Length, $(if ($x.Contains("`r`n")) { 'CR LF' } elseif ($x.Contains("`n")) { 'LF' } else { 'one line' })) }
        Check "$id.X" ($k -eq $(if ($exp.customXml) { 1 } else { 0 })) $det
    } catch { Check "$id.X" $false ("CustomXMLParts: " + (Msg $_)) }
}

Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$utf8 = New-Object System.Text.UTF8Encoding($false)
function Read-Entry($e) { $r = New-Object System.IO.StreamReader($e.Open(), $utf8); try { $r.ReadToEnd() } finally { $r.Close() } }
function Count-Of($text, $what) { $n = 0; $i = 0; while (($i = $text.IndexOf($what, $i)) -ge 0) { $n++; $i += $what.Length }; $n }

# What Excel wrote (the Mac checker, check_m3_winsaved.mjs, goes further).
function Inspect-Saved($path, $id) {
    try {
        $z = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            $wbx = Read-Entry ($z.GetEntry('xl/workbook.xml'))
            $cp = $wbx.IndexOf('<calcPr'); $calc = if ($cp -ge 0) { $wbx.Substring($cp, [Math]::Min(80, $wbx.Length - $cp)).Split('>')[0] + '>' } else { 'none' }
            Add-Line "$id.Z" 'INFO' ("saved workbook.xml: {0} x '_x000a_', {1} x '&#10;' in names; {2}; calcChain.xml {3}" -f (Count-Of $wbx '_x000a_'), (Count-Of $wbx '&#10;'), $calc, $(if ($z.GetEntry('xl/calcChain.xml')) { 'present' } else { 'absent' }))
            foreach ($e in $z.Entries) {
                if ($e.FullName -like 'customXml/item*.xml') {
                    $t = Read-Entry $e
                    if ($t.Contains($kit.embedNamespace)) { Add-Line "$id.Z" 'INFO' ("saved {0}: {1} bytes; line ends: {2} CR LF of {3}" -f $e.FullName, $e.Length, (Count-Of $t "`r`n"), (Count-Of $t "`n")) }
                }
                if ($e.FullName -like 'customXml/itemProps*.xml') { $t = Read-Entry $e; if ($t.Contains($kit.embedNamespace)) { Add-Line "$id.Z" 'INFO' ("saved {0} keeps the schemaRef" -f $e.FullName) } }
            }
        } finally { $z.Dispose() }
    } catch { Add-Line "$id.Z" 'WARN' ("cannot read the saved copy: " + (Msg $_)) }
}

function Save-Copy($wb, $name, $id) {
    $dest = Join-Path $winDir ("{0}_winsaved.xlsx" -f $name)
    try {
        if (Test-Path $dest) { Remove-Item $dest -Force }
        $wb.SaveAs($dest, 51)
        Check "$id.W" (Test-Path $dest) ("saved " + $dest)
        return $dest
    } catch { Check "$id.W" $false ("SaveAs: " + (Msg $_)); return $null }
}

function New-Excel {
    $before = Excel-Pids
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true
    $xl.DisplayAlerts = $true   # a repair prompt must show, not be answered for us
    $mine = @(Excel-Pids | Where-Object { $before -notcontains $_ })
    return @{ xl = $xl; pids = $mine }
}
function Quit-Excel($h) {
    try { $h.xl.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($h.xl) } catch {}
    $h.xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    # COM keeps an instance alive while PowerShell holds references to its objects; give it
    # time, then end only the process this script started.
    for ($i = 0; $i -lt 20; $i++) { if (@(Excel-Pids | Where-Object { $h.pids -contains $_ }).Count -eq 0) { return }; Start-Sleep -Milliseconds 500 }
    foreach ($p in $h.pids) { try { Stop-Process -Id $p -Force; Add-Line 'Q00' 'INFO' "ended the Excel process this script started (pid $p), which did not quit" } catch {} }
}

# ---- part 1: open every kit workbook --------------------------------------------------------
if ($ReopenOnly) { Add-Line 'E06' 'INFO' 'part 1 skipped (-ReopenOnly)' } else {
    Say "Part 1: opening the kit workbooks in a new Excel window. Do not click in it until asked."
    $clean = @($kit.files | Where-Object { $_.repair -eq $false })
    $risky = @($kit.files | Where-Object { $_.repair -ne $false })
    $h = $null
    try {
        $h = New-Excel
        Add-Line 'T00' 'INFO' ("Excel {0} build {1}; decimal sep '{2}' list sep '{3}'; pid {4}" -f $h.xl.Version, $h.xl.Build, $h.xl.International(3), $h.xl.International(5), ($h.pids -join ','))
    } catch { Add-Line 'T00' 'FAIL' ("cannot start Excel through COM: " + (Msg $_)); Save-Report; return }

    $i = 0
    foreach ($f in $clean) {
        $i++
        $id = 'K{0:d2}' -f $i
        $name = [System.IO.Path]::GetFileNameWithoutExtension($f.file)
        $exp = Get-Content -Raw -Encoding UTF8 (Join-Path $kitDir $f.expected) | ConvertFrom-Json
        Add-Line $id 'INFO' ("== {0}: {1}" -f $f.file, $exp.title)
        $path = Stage-Copy (Join-Path $kitDir $f.file) $f.file
        $t0 = Get-Date
        Start-Dog 20 $dialogText
        $wb = $null; $err = ''
        try { $wb = $h.xl.Workbooks.Open($path) } catch { $err = Msg $_ } finally { Stop-Dog }
        $secs = ((Get-Date) - $t0).TotalSeconds
        if ($null -eq $wb) { Check "$id.O" $false ("did not open ({0:n1} s): {1}" -f $secs, $err); Save-Report; continue }
        $cap = try { $h.xl.ActiveWindow.Caption } catch { '?' }
        $saw = ''
        if ($secs -gt 15) { $saw = Read-Host "Opening $($f.file) took $([int]$secs) s. Did Excel show a dialog? Type what it said (Enter = no dialog)" }
        Check "$id.O" ($saw -eq '') ("opened in {0:n1} s; caption '{1}'{2}" -f $secs, $cap, $(if ($saw) { "; the author saw: $saw" } else { '' }))
        Read-Kit $h.xl $wb $exp $id
        $dest = Save-Copy $wb $name $id
        try { $wb.Close($false) } catch { Add-Line "$id.Q" 'WARN' ("close: " + (Msg $_)) }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {}
        $wb = $null
        if ($dest) { Inspect-Saved $dest $id }
        Save-Report
    }
    Quit-Excel $h

    foreach ($f in $risky) {
        $i++
        $id = 'K{0:d2}' -f $i
        $name = [System.IO.Path]::GetFileNameWithoutExtension($f.file)
        $exp = Get-Content -Raw -Encoding UTF8 (Join-Path $kitDir $f.expected) | ConvertFrom-Json
        Add-Line $id 'INFO' ("== {0}: {1}" -f $f.file, $exp.title)
        Say ("Next: {0}. Expected: Excel {1}." -f $f.file, $(if ($f.repair -eq $true) { 'asks to repair it' } else { 'refuses it or asks to repair it' }))
        Write-Host $exp.repairHint -ForegroundColor Yellow
        Write-Host "A new Excel window opens for this file alone. Press Enter here to start." -ForegroundColor Yellow
        [void](Read-Host)
        $hx = $null
        try { $hx = New-Excel } catch { Add-Line "$id.O" 'FAIL' ("cannot start Excel: " + (Msg $_)); continue }
        $t0 = Get-Date
        Start-Dog 25 $dialogText
        $wb = $null; $err = ''
        try { $wb = $hx.xl.Workbooks.Open((Stage-Copy (Join-Path $kitDir $f.file) $f.file)) } catch { $err = Msg $_ } finally { Stop-Dog }
        $secs = ((Get-Date) - $t0).TotalSeconds
        $cap = if ($wb) { try { $hx.xl.ActiveWindow.Caption } catch { '?' } } else { '' }
        $q = Read-Host "Did Excel ask whether to repair (recover) $($f.file)? [y/n]"
        $asked = $q -match '^\s*[yYsS]'
        $saw = Read-Host "What else did Excel show (messages, list of repairs)? Type it in short, or Enter"
        $state = if ($wb) { 'opened' } else { "not opened: $err" }
        Add-Line "$id.O" 'INFO' ("{0} after {1:n1} s; caption '{2}'; repair prompt: {3}; the author saw: {4}" -f $state, $secs, $cap, $(if ($asked) { 'yes' } else { 'no' }), $(if ($saw) { $saw } else { '-' }))
        if ($f.repair -eq $true) { Check "$id.R" $asked ("repair prompt shown, as on the Mac (F8 p2: stale calcChain)") }
        else { Add-Line "$id.R" ($(if ($asked -or -not $wb) { 'PASS' } else { 'WARN' })) ($(if (-not $wb) { 'Excel refused the file, as on the Mac' } elseif ($asked) { 'Excel offered to repair it (the Mac refused it outright)' } else { 'Excel opened it with no prompt: Windows accepts a 256-character comment' })) }
        if ($wb) {
            Read-Kit $hx.xl $wb $exp $id
            $dest = Save-Copy $wb $name $id
            try { $wb.Close($false) } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {}
            $wb = $null
            if ($dest) { Inspect-Saved $dest $id }
        }
        Quit-Excel $hx
        Save-Report
    }
}

# ---- part 2: the --reopen script (E7) ------------------------------------------------------
# Runs m3kit\xln_excel.ps1 exactly as the CLI does (packages/cli/src/excel.ts): copied to a
# temp file, `powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File ...`,
# same time limits (state 20 s, close 30 s, open 90 s).
$reo = Join-Path $stage 'reopen'
$script:xlnPs1 = $null
function Invoke-Xln($action, $path, $timeoutSec) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Action {1} -Path "{2}"' -f $script:xlnPs1, $action, $path
    $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.CreateNoWindow = $true
    $t0 = Get-Date
    $p = [System.Diagnostics.Process]::Start($psi)
    $so = $p.StandardOutput.ReadToEndAsync(); $se = $p.StandardError.ReadToEndAsync()
    $done = $p.WaitForExit($timeoutSec * 1000)
    if (-not $done) { try { $p.Kill() } catch {} }
    [void]$so.Wait(5000); [void]$se.Wait(5000)
    $r = [pscustomobject]@{
        out = $(if ($so.IsCompleted) { $so.Result.Trim() } else { '' })
        err = $(if ($se.IsCompleted) { $se.Result.Trim() } else { '' })
        code = $(if ($done) { $p.ExitCode } else { $null })
        timedOut = -not $done
        secs = ((Get-Date) - $t0).TotalSeconds
    }
    return $r
}
function Show-Run($r) {
    $s = "'{0}' in {1:n1} s" -f $r.out, $r.secs
    if ($r.timedOut) { $s += ' TIMED OUT (killed)' } elseif ($r.code -ne 0) { $s += " exit $($r.code)" }
    if ($r.err) { $s += '; stderr: ' + (Short $r.err 300) }
    return $s
}
# The workbook as the running Excel lists it, by file name (not by path: a OneDrive copy
# reports a URL as FullName).
# Also through the Running Object Table, as the --reopen script does: a workbook opened from
# Explorer can live in another Excel process (measured in the first -ReopenOnly run).
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
}
"@
function Get-RunningBook($fileName, $path) {
    try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $app = $null }
    if ($app) { foreach ($b in $app.Workbooks) { if ($b.Name -ieq $fileName) { return $b } } }
    if ($path) { $o = [XlnRot]::Find([System.IO.Path]::GetFullPath($path)); if ($o -ne $null) { return $o } }
    return $null
}
# Polls 'state' until the script finds the workbook open ('saved' or 'unsaved').
function Wait-Open($path, $seconds) {
    $t0 = Get-Date; $r = $null
    while (((Get-Date) - $t0).TotalSeconds -lt $seconds) {
        $r = Invoke-Xln 'state' $path 20
        if ($r.out -eq 'saved' -or $r.out -eq 'unsaved') { break }
        Start-Sleep -Seconds 2
    }
    return $r
}
$dirtyNote = "Excel reports unsaved changes in a workbook nobody edited: it marks a workbook changed after recalculating it on load (fullCalcOnLoad, or a calc engine newer than the file's calcId), so the next 'xln build --reopen' would refuse (exit 5) without --discard"

if ($doReopen) {
    Say "Part 2: the PowerShell script behind 'xln build --reopen'. Excel will open and close by itself a few times."
    $left = Excel-Pids
    for ($k = 0; $k -lt 20 -and $left.Count -gt 0; $k++) { Start-Sleep -Milliseconds 500; $left = Excel-Pids }
    if ($left.Count -gt 0) {
        Write-Host "Excel is still running (pid $($left -join ', ')). Close every Excel window, then press Enter (or S and Enter to skip part 2)." -ForegroundColor Yellow
        $a = Read-Host
        if ($a -match '^\s*[sS]') { $doReopen = $false; Add-Line 'R00' 'INFO' 'part 2 skipped: Excel was running' }
        elseif ((Excel-Pids).Count -gt 0) { $doReopen = $false; Add-Line 'R00' 'WARN' 'part 2 skipped: Excel still running' }
    }
}
if ($doReopen) {
    try {
        if (Test-Path $reo) { Remove-Item $reo -Recurse -Force }
        New-Item -ItemType Directory $reo -Force | Out-Null
        $tmpPs1 = Join-Path $reo 'excel.ps1'
        Copy-Item (Join-Path $kitDir $kit.reopen.script) $tmpPs1
        try { Unblock-File $tmpPs1 } catch {}
        $script:xlnPs1 = $tmpPs1
        $book = Join-Path $reo 'reopen_check.xlsx'
        Copy-Item (Join-Path $kitDir $kit.reopen.workbook) $book
        try { Unblock-File $book } catch {}
        $bookName = Split-Path $book -Leaf
        Add-Line 'R00' 'INFO' ("script {0}; workbook {1} (a copy of {2})" -f $tmpPs1, $book, $kit.reopen.workbook)
        $hash0 = (Get-FileHash $book -Algorithm SHA256).Hash

        $r = Invoke-Xln 'state' $book 20
        Check 'R01' ($r.out -eq 'notrunning') ("state with Excel not running: " + (Show-Run $r) + " (expected 'notrunning')")

        $r = Invoke-Xln 'open' $book 90
        Check 'R02' ($r.out -eq 'opened') ("open with Excel not running (the script starts Excel): " + (Show-Run $r) + " (expected 'opened')")
        Start-Sleep -Seconds 2
        $alive = (Excel-Pids).Count -gt 0
        Check 'R03' $alive ("Excel still running after the script ended: {0}" -f $alive)

        $r = Invoke-Xln 'state' $book 20
        if ($r.out -eq 'notrunning') {
            Add-Line 'R04' 'WARN' ("state right after the script opened it: " + (Show-Run $r) + ": Excel is running but not found by GetActiveObject (not in the Running Object Table yet?)")
            Write-Host "Click once on this console window (so Excel loses focus), then press Enter." -ForegroundColor Yellow
            [void](Read-Host)
            $r = Invoke-Xln 'state' $book 20
            Check 'R04' ($r.out -eq 'saved') ("state after Excel lost focus: " + (Show-Run $r) + " (expected 'saved')")
        } elseif ($r.out -eq 'unsaved') { Add-Line 'R04' 'WARN' ("state, just opened: " + (Show-Run $r) + ": " + $dirtyNote) }
        else { Check 'R04' ($r.out -eq 'saved') ("state, open and saved: " + (Show-Run $r) + " (expected 'saved')") }

        $r = Invoke-Xln 'close' $book 30
        Check 'R05' ($r.out -eq 'closed') ("close: " + (Show-Run $r) + " (expected 'closed')")
        $r = Invoke-Xln 'state' $book 20
        Check 'R06' ($r.out -eq 'closed') ("state after close, Excel still running: " + (Show-Run $r) + " (expected 'closed')")
        $r = Invoke-Xln 'close' $book 30
        Check 'R07' ($r.out -eq 'not open') ("close when not open: " + (Show-Run $r) + " (expected 'not open')")

        # The author's way: the file opened from Explorer (here: Start-Process, the same association).
        Start-Process -FilePath $book
        $r = Wait-Open $book 30
        if ($r.out -ne 'saved' -and $r.out -ne 'unsaved') {
            Write-Host "Excel opened $bookName but the script does not see it yet. Click once on this console window, then press Enter." -ForegroundColor Yellow
            [void](Read-Host)
            $r2 = Wait-Open $book 20
            Check 'R08' ($r2.out -eq 'saved') ("state of a workbook opened from Explorer: first " + (Show-Run $r) + "; after clicking this window " + (Show-Run $r2))
        } elseif ($r.out -eq 'unsaved') { Add-Line 'R08' 'WARN' ("state of a workbook opened from Explorer: " + (Show-Run $r) + ": " + $dirtyNote) }
        else { Check 'R08' $true ("state of a workbook opened from Explorer: " + (Show-Run $r)) }
        $b = Get-RunningBook $bookName $book
        if ($b) { Add-Line 'R08' 'INFO' ("Excel reports FullName '{0}' for {1}" -f $b.FullName, $book) }

        # Unsaved change: through COM, or by the author if COM cannot reach the workbook.
        $changed = $false
        if ($b) { try { $b.Worksheets.Item(1).Range('Z99').Value2 = 1; $changed = $true } catch { Add-Line 'R09' 'INFO' ("COM edit failed: " + (Msg $_)) } }
        if (-not $changed) {
            Write-Host "In Excel, type any number into an empty cell of $bookName and press Enter (do NOT save). Then press Enter here." -ForegroundColor Yellow
            [void](Read-Host)
        }
        $r = Invoke-Xln 'state' $book 20
        Check 'R09' ($r.out -eq 'unsaved') ("state with an unsaved change: " + (Show-Run $r) + " (expected 'unsaved'; xln then refuses with exit 5 unless --discard)")
        $r = Invoke-Xln 'close' $book 30
        Check 'R10' ($r.out -eq 'closed') ("close (discarding the change, as --discard): " + (Show-Run $r) + " (expected 'closed')")
        Start-Sleep -Seconds 1
        $hash1 = (Get-FileHash $book -Algorithm SHA256).Hash
        Check 'R11' ($hash1 -eq $hash0) ("the file on disk is unchanged after close without saving: {0}" -f ($hash1 -eq $hash0))

        # What a build does between close and open: write a different workbook in place.
        Copy-Item (Join-Path $kitDir $kit.reopen.rebuilt) $book -Force
        $r = Invoke-Xln 'open' $book 90
        Check 'R12' ($r.out -eq 'opened') ("open the rewritten file, Excel running: " + (Show-Run $r) + " (expected 'opened')")
        $b = Get-RunningBook $bookName $book
        if ($b) {
            $cx = try { $b.CustomXMLParts.SelectByNamespace($kit.embedNamespace).Count } catch { -1 }
            Check 'R13' ($cx -eq 1) ("the reopened workbook is the new file (embedded part present: {0})" -f $cx)
        } else { Add-Line 'R13' 'WARN' 'cannot reach the reopened workbook through COM' }
        $r = Invoke-Xln 'state' $book 20
        if ($r.out -eq 'unsaved') { Add-Line 'R14' 'WARN' ("state after reopening the built file: " + (Show-Run $r) + ": " + $dirtyNote) } else { Check 'R14' ($r.out -eq 'saved') ("state after reopen: " + (Show-Run $r) + " (expected 'saved')") }

        # A file that needs repair, opened by the script (CorruptLoad = xlNormalLoad).
        $bad = Join-Path $reo 'reopen_repair.xlsx'
        Copy-Item (Join-Path $kitDir $kit.reopen.repair) $bad
        try { Unblock-File $bad } catch {}
        Say "Next: the script opens reopen_repair.xlsx, which needs repair."
        Write-Host "If Excel asks whether to recover its contents, click No [No] within 60 seconds" -ForegroundColor Yellow
        Write-Host "(that is what xln tells its users to do). Press Enter here to start." -ForegroundColor Yellow
        [void](Read-Host)
        $r = Invoke-Xln 'open' $bad 90
        $q = Read-Host "Did Excel ask whether to repair reopen_repair.xlsx? [y/n]"
        $asked = $q -match '^\s*[yYsS]'
        Add-Line 'R15' 'INFO' ("open of a file needing repair: " + (Show-Run $r) + "; repair prompt: " + $(if ($asked) { 'yes' } else { 'no' }))
        Check 'R15' ($r.out -ne 'opened') ("xln must not report 'opened' for a file it could not open cleanly: script said '{0}'" -f $r.out)
        if ($r.timedOut) {
            Write-Host "The script timed out: if a dialog is still open in Excel, answer it (No), then press Enter." -ForegroundColor Yellow
            [void](Read-Host)
        }
        $r = Invoke-Xln 'state' $bad 20
        Add-Line 'R16' 'INFO' ("state of the file needing repair afterwards: " + (Show-Run $r))

        # Optional: a workbook in the OneDrive folder, where Excel may report a URL as its
        # FullName and the script, which compares paths, would not find it.
        if ($env:OneDrive -and (Test-Path $env:OneDrive)) {
            $a = Read-Host "Also try a copy in your OneDrive folder ($env:OneDrive)? It is deleted afterwards. [y/n]"
            if ($a -match '^\s*[yYsS]') {
                $od = Join-Path $env:OneDrive 'xln-m3check-onedrive.xlsx'
                try {
                    Copy-Item (Join-Path $kitDir $kit.reopen.workbook) $od -Force
                    Start-Process -FilePath $od
                    $r = Wait-Open $od 40
                    $ob = Get-RunningBook (Split-Path $od -Leaf)
                    $fn = if ($ob) { $ob.FullName } else { '(not reachable through COM)' }
                    Check 'R17' ($r.out -eq 'saved' -or $r.out -eq 'unsaved') ("state of a workbook in OneDrive: " + (Show-Run $r) + "; Excel's FullName '$fn'")
                    $r = Invoke-Xln 'close' $od 30
                    Check 'R18' ($r.out -eq 'closed') ("close of the OneDrive workbook: " + (Show-Run $r))
                    if ($ob) { try { $ob.Close($false) } catch {} }
                } catch { Add-Line 'R17' 'WARN' ("OneDrive test: " + (Msg $_)) }
                finally { $ob = $null; Start-Sleep -Seconds 2; try { Remove-Item $od -Force } catch { Add-Line 'R18' 'INFO' "could not delete $od (delete it yourself)" } }
            }
        }
    } catch {
        Add-Line 'R99' 'FAIL' ("part 2 stopped: " + (Msg $_))
    } finally {
        # Close our two workbooks without saving, then quit Excel if nothing else is open
        # (it was started by this check: Excel was not running when part 2 began).
        try {
            $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
            foreach ($nm in 'reopen_check.xlsx', 'reopen_repair.xlsx', 'xln-m3check-onedrive.xlsx') { foreach ($b in @($app.Workbooks)) { if ($b.Name -ieq $nm) { $b.Close($false) } } }
            if ($app.Workbooks.Count -eq 0) { $app.Quit() }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
        } catch {}
        $app = $null; $b = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        $left = @()
        for ($k = 0; $k -lt 10; $k++) { Start-Sleep -Seconds 1; $left = @(Excel-Pids | Where-Object { $startPids -notcontains $_ }); if ($left.Count -eq 0) { break } }
        # Excel was not running when part 2 began, so a windowless Excel left now is one
        # this check started that did not exit after Quit (COM references): end it.
        foreach ($p in $left) {
            $proc = Get-Process -Id $p -ErrorAction SilentlyContinue
            if ($proc -and $proc.MainWindowHandle -eq [IntPtr]::Zero) { try { Stop-Process -Id $p -Force; Add-Line 'R98' 'INFO' "ended a windowless Excel left by part 2 (pid $p)" } catch {} }
        }
        if (@(Excel-Pids | Where-Object { $startPids -notcontains $_ }).Count -gt 0) { Write-Host "Excel is still open from part 2: close it yourself (do not save reopen_*.xlsx)." -ForegroundColor Yellow }
        Save-Report
    }
}

# ---- summary --------------------------------------------------------------------------------
Add-Line 'S00' 'INFO' ("summary: {0} PASS, {1} FAIL, {2} WARN" -f $script:count.PASS, $script:count.FAIL, $script:count.WARN)
foreach ($l in $script:fails) { $script:rpt.Add("S01 | LIST | $l") }
Save-Report
Write-Host ''
Write-Host ("report: {0}" -f $outFile)
Write-Host ("saved copies: {0}" -f $winDir)
Write-Host ''
Write-Host 'Now send the results back. In a terminal (or Git Bash) in the repo folder:' -ForegroundColor Cyan
Write-Host '    git add probes/results'
Write-Host '    git commit -m "M3 Windows check results"'
Write-Host '    git push'
Write-Host '(If push is refused, run "git pull --rebase" first, then "git push" again.)'
