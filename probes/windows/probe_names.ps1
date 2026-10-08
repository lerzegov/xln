# probe_names.ps1 -- P0 probe: how much of the Name Manager can PowerShell drive via COM?
#
#   run_windows.cmd                       (double-click; uses -ExecutionPolicy Bypass)
#   powershell -NoProfile -ExecutionPolicy Bypass -File probe_names.ps1
#   ...or, if scripts are blocked by policy: open PowerShell, paste this whole file.
#
# Starts its OWN Excel instance, works on a NEW blank workbook, saves a copy to
# %TEMP%\excel-dim-probe so the stored form can be read from xl/workbook.xml, closes it
# unsaved and quits that instance. Your other open workbooks are not touched.
# Test IDs match probes/mac/probe_names.applescript and probes/officescripts/probe_names.ts.
# Output: "Txx | STATUS | detail" lines, written to probes\results\win-<host>-<time>.txt.

$ErrorActionPreference = 'Stop'
$script:rpt = New-Object System.Collections.Generic.List[string]
function Add-Line($tid, $verdict, $det) {
    $line = "$tid | $verdict | " + ("$det" -replace '\s*[\r\n]+\s*', ' / ')
    $script:rpt.Add($line); Write-Host $line
}
function Short($s, $n = 80) { if ($null -eq $s) { return '<null>' }; $t = "$s"; if ($t.Length -gt $n) { $t.Substring(0, $n) + '...' } else { $t } }

$here = if ($PSScriptRoot) { $PSScriptRoot } else { Join-Path $env:USERPROFILE 'Desktop' }
$outDir = Join-Path $here '..\results'
if (-not (Test-Path $outDir)) { try { New-Item -ItemType Directory $outDir | Out-Null } catch { $outDir = $env:TEMP } }
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$outFile = Join-Path $outDir ("win-{0}-{1}.txt" -f $env:COMPUTERNAME, $stamp)
$stage = Join-Path $env:TEMP 'excel-dim-probe'
if (-not (Test-Path $stage)) { New-Item -ItemType Directory $stage | Out-Null }
$saved = Join-Path $stage 'probe_win.xlsx'
if (Test-Path $saved) { Remove-Item $saved -Force }

Add-Line 'H00' 'INFO' ("# excel-dim P0 probe -- Windows  script v4 (T22 routes + file-level F)  {0:u}" -f (Get-Date).ToUniversalTime())
# --- environment: the things a managed PC can forbid ---------------------------------
Add-Line 'E01' 'INFO' ("OS {0}; PowerShell {1}; culture {2}" -f [Environment]::OSVersion.VersionString, $PSVersionTable.PSVersion, (Get-Culture).Name)
$lm = $ExecutionContext.SessionState.LanguageMode
Add-Line 'E02' ($(if ($lm -eq 'FullLanguage') { 'PASS' } else { 'WARN' })) "LanguageMode=$lm (ConstrainedLanguage blocks most COM objects)"
try { Get-ExecutionPolicy -List | ForEach-Object { Add-Line 'E03' 'INFO' ("ExecutionPolicy {0}={1}" -f $_.Scope, $_.ExecutionPolicy) } } catch { Add-Line 'E03' 'FAIL' $_.Exception.Message }

try {
    $xl = New-Object -ComObject Excel.Application
    Add-Line 'E04' 'PASS' 'COM Excel.Application created'
} catch {
    Add-Line 'E04' 'FAIL' ("cannot create Excel COM object: " + $_.Exception.Message)
    $script:rpt | Set-Content -Encoding UTF8 $outFile; Write-Host "report: $outFile"; return
}

$M = [System.Reflection.Missing]::Value
# Run 2 (2026-10-04) showed that Names.Add(name, refersTo) called from PowerShell parses
# refersTo in the LOCAL language (on it-IT, '=LAMBDA(x, x+1)' is rejected), whatever the
# thread culture, while property puts (.Formula2, .RefersTo) take English. T22 tries four
# ways to create a name; Add-Name then uses the first English route that works.
$native = [System.Threading.Thread]::CurrentThread.CurrentCulture
$enUS = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
$script:route = 'a'
function Invoke-Add($coll, $nm, $f, $culture) {
    [System.__ComObject].InvokeMember('Add', [System.Reflection.BindingFlags]::InvokeMethod, $null, $coll, [object[]]@($nm, $f), $culture)
}
function Add-Name($coll, $nm, $f) {
    switch ($script:route) {
        'c' { return (Invoke-Add $coll $nm $f $enUS) }
        'b' {
            $n = $coll.Add($nm, '=0')
            try { $n.RefersTo = $f } catch { try { $n.Delete() } catch {}; throw }
            return $n
        }
        default { return $coll.Add($nm, $f) }
    }
}
try {
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    Add-Line 'T00' 'INFO' ("Excel {0} build {1}; decimal sep '{2}' list sep '{3}'" -f $xl.Version, $xl.Build, $xl.International(3), $xl.International(5))

    $wb = $xl.Workbooks.Add()
    $s1 = $wb.Worksheets.Item(1); $s1.Name = 'S1'
    $s2 = $wb.Worksheets.Add($M, $s1); $s2.Name = 'S2'
    $s1.Activate() | Out-Null

    # T22 how to create a name with an ENGLISH definition from PowerShell (see header)
    $f22 = '=LAMBDA(x, x+1)'
    $routes = [ordered]@{
        'a' = @('positional Names.Add', { param($nm) $wb.Names.Add($nm, $f22) })
        'b' = @('Names.Add then .RefersTo put', { param($nm) $n = $wb.Names.Add($nm, '=0'); $n.RefersTo = $f22; $n })
        'c' = @('InvokeMember Add, en-US', { param($nm) Invoke-Add $wb.Names $nm $f22 $enUS })
        'd' = @(("InvokeMember Add, " + $native.Name), { param($nm) Invoke-Add $wb.Names $nm $f22 $native })
    }
    $ok = @{}
    foreach ($k in $routes.Keys) {
        $nm = "P_Route_$k"
        try {
            $n = & $routes[$k][1] $nm
            $s1.Range('A20').Formula2 = "=$nm(41)"; $xl.Calculate()
            $ok[$k] = ($s1.Range('A20').Value2 -eq 42)
            Add-Line 'T22' ($(if ($ok[$k]) { 'PASS' } else { 'FAIL' })) ("{0} {1}: RefersTo={2} -> {3}" -f $k, $routes[$k][0], $n.RefersTo, $s1.Range('A20').Text)
        } catch { $ok[$k] = $false; Add-Line 'T22' 'FAIL' ("{0} {1}: {2}" -f $k, $routes[$k][0], (Short $_.Exception.GetBaseException().Message)) }
        try { $wb.Names.Item($nm).Delete() } catch {}
    }
    $s1.Range('A20').ClearContents() | Out-Null
    foreach ($k in 'c', 'b', 'a') { if ($ok[$k]) { $script:route = $k; break } }
    Add-Line 'T22' 'INFO' ("tests below create names via route " + $script:route)

    # T01 workbook-scoped constant
    try { $n = (Add-Name $wb.Names 'P_K' '=10'); Add-Line 'T01' 'PASS' ("created P_K; RefersTo=" + $n.RefersTo) } catch { Add-Line 'T01' 'FAIL' $_.Exception.Message }

    # T02 LAMBDA name called from a cell
    try {
        (Add-Name $wb.Names 'P_Add1' '=LAMBDA(x, x+1)') | Out-Null
        $s1.Range('A1').Formula2 = '=P_Add1(41)'; $xl.Calculate()
        $v = $s1.Range('A1').Value2
        Add-Line 'T02' ($(if ($v -eq 42) { 'PASS' } else { 'FAIL' })) ("=P_Add1(41) -> " + $s1.Range('A1').Text)
    } catch { Add-Line 'T02' 'FAIL' $_.Exception.Message }

    # T03 stored text, English vs local
    try {
        $n = $wb.Names.Item('P_Add1')
        Add-Line 'T03' 'INFO' ("RefersTo=" + $n.RefersTo)
        Add-Line 'T03' 'INFO' ("RefersToLocal=" + $n.RefersToLocal)
    } catch { Add-Line 'T03' 'FAIL' $_.Exception.Message }

    # T04 Name Manager comment
    try {
        $wb.Names.Item('P_Add1').Comment = 'probe comment'
        Add-Line 'T04' 'PASS' ("comment read back: " + $wb.Names.Item('P_Add1').Comment)
    } catch { Add-Line 'T04' 'FAIL' ("comment: " + $_.Exception.Message) }

    # T05 sheet-scoped names
    try {
        (Add-Name $s2.Names 'P_Local' '=5') | Out-Null
        $s2.Range('A1').Formula2 = '=P_Local'; $s1.Range('A3').Formula2 = '=S2!P_Local'; $xl.Calculate()
        Add-Line 'T05a' 'INFO' ("make at sheet: S2!A1=" + $s2.Range('A1').Text + " S1!A3=" + $s1.Range('A3').Text)
    } catch { Add-Line 'T05a' 'FAIL' $_.Exception.Message }
    try { (Add-Name $wb.Names 'S2!P_Local2' '=6') | Out-Null; Add-Line 'T05b' 'INFO' "make at wb with 'S2!' prefix: ok" } catch { Add-Line 'T05b' 'FAIL' $_.Exception.Message }
    try { $l = @(); foreach ($x in $s2.Names) { $l += $x.Name }; Add-Line 'T05c' 'INFO' ("names of S2: " + ($l -join '; ')) } catch { Add-Line 'T05c' 'FAIL' $_.Exception.Message }

    # T06 name over a spilled range
    try {
        $s1.Range('B1').Formula2 = '=SEQUENCE(1,5)'
        (Add-Name $wb.Names 'P_Spill' '=S1!$B$1#') | Out-Null
        $s1.Range('C2').Formula2 = '=COLUMNS(P_Spill)'; $xl.Calculate()
        Add-Line 'T06' 'INFO' ("RefersTo=" + $wb.Names.Item('P_Spill').RefersTo + "  COLUMNS->" + $s1.Range('C2').Text)
    } catch { Add-Line 'T06' 'FAIL' $_.Exception.Message }

    # T08 update a definition
    try {
        $s1.Range('A4').Formula2 = '=P_K'
        $wb.Names.Item('P_K').RefersTo = '=20'; $xl.Calculate()
        Add-Line 'T08' 'INFO' ("after update S1!A4=" + $s1.Range('A4').Text + " (expect 20)")
    } catch { Add-Line 'T08' 'FAIL' $_.Exception.Message }

    # T09 rename: do dependents follow?
    try { $wb.Names.Item('P_K').Name = 'P_K2'; Add-Line 'T09' 'INFO' ("renamed; S1!A4 formula now " + $s1.Range('A4').Formula) } catch { Add-Line 'T09' 'FAIL' $_.Exception.Message }

    # T10 delete
    try { $wb.Names.Item('P_K2').Delete(); $xl.Calculate(); Add-Line 'T10' 'INFO' ("deleted; S1!A4 shows " + $s1.Range('A4').Text) } catch { Add-Line 'T10' 'FAIL' $_.Exception.Message }

    # T11 syntax error: what does the caller get?
    try { $n = (Add-Name $wb.Names 'P_Bad' '=LAMBDA(x, x+'); Add-Line 'T11' 'INFO' ("accepted (!) RefersTo=" + $n.RefersTo) } catch { Add-Line 'T11' 'INFO' ("rejected with: " + $_.Exception.Message) }

    # T12 collision with a built-in
    try {
        (Add-Name $wb.Names 'Fact' '=LAMBDA(n, 1)') | Out-Null
        $s1.Range('A5').Formula2 = '=Fact(5)'; $xl.Calculate()
        Add-Line 'T12' 'INFO' ("=Fact(5) -> " + $s1.Range('A5').Text + " (1 = name wins, 120 = built-in wins); formula reads " + $s1.Range('A5').Formula)
    } catch { Add-Line 'T12' 'INFO' ("rejected: " + $_.Exception.Message) }

    # T13 dotted (AFE-module-style) name
    try {
        (Add-Name $wb.Names 'Mod.Fn' '=LAMBDA(x, x*10)') | Out-Null
        $s1.Range('A6').Formula2 = '=Mod.Fn(2)'; $xl.Calculate()
        Add-Line 'T13' 'INFO' ("=Mod.Fn(2) -> " + $s1.Range('A6').Text)
    } catch { Add-Line 'T13' 'FAIL' $_.Exception.Message }

    # T14 line breaks inside a definition
    try {
        (Add-Name $wb.Names 'P_Multi' "=LAMBDA(x,`n  x*2)") | Out-Null
        $r = $wb.Names.Item('P_Multi').RefersTo
        $s1.Range('A7').Formula2 = '=P_Multi(4)'; $xl.Calculate()
        $keep = if ($r -match "`n") { 'newline KEPT' } else { 'newline LOST' }
        Add-Line 'T14' 'INFO' ("$keep; =P_Multi(4) -> " + $s1.Range('A7').Text)
    } catch { Add-Line 'T14' 'FAIL' $_.Exception.Message }

    # T15 length limit (same lengths as the Mac probe: 7992, 8192, 9012)
    foreach ($k in 3990, 4090, 4500) {
        $f = '=LAMBDA(x,x' + ('+1' * $k) + ')'; $nm = "P_Long$k"
        try {
            (Add-Name $wb.Names $nm $f) | Out-Null
            $s1.Range('A9').Formula2 = "=$nm(0)"; $xl.Calculate()
            Add-Line 'T15' 'INFO' ("{0} chars accepted; ->{1}" -f $f.Length, $s1.Range('A9').Text)
        } catch { Add-Line 'T15' 'INFO' ("{0} chars rejected: {1}" -f $f.Length, $_.Exception.Message) }
    }

    # T16 evaluate an expression without touching a cell
    try { Add-Line 'T16' 'INFO' ("evaluate P_Add1(1) -> " + $xl.Evaluate('P_Add1(1)')) } catch { Add-Line 'T16' 'FAIL' $_.Exception.Message }
    try { Add-Line 'T16' 'INFO' ("evaluate ROWS*10+COLUMNS of P_Spill -> " + $xl.Evaluate('ROWS(P_Spill)*10+COLUMNS(P_Spill)')) } catch { Add-Line 'T16' 'FAIL' $_.Exception.Message }

    # T17 non-ASCII name (Grow + Greek small lambda)
    try {
        $g = 'Grow' + [char]0x03BB
        (Add-Name $wb.Names $g '=LAMBDA(b, g, b*(1+g))') | Out-Null
        $s1.Range('A10').Formula2 = "=$g(100, 0.1)"; $xl.Calculate()
        Add-Line 'T17' 'INFO' ("=Grow<lambda>(100,0.1) -> " + $s1.Range('A10').Text)
    } catch { Add-Line 'T17' 'FAIL' $_.Exception.Message }

    # T18 write in the LOCAL language (list separator from T00): positional RefersToLocal
    try {
        $sep = $xl.International(5)
        $n = $wb.Names.Add('P_Loc', $M, $true, $M, $M, $M, $M, "=LAMBDA(x$sep x+1)")
        Add-Line 'T18' 'INFO' ("RefersToLocal with '$sep' accepted; RefersTo=" + $n.RefersTo + " RefersToLocal=" + $n.RefersToLocal)
    } catch { Add-Line 'T18' 'INFO' ("RefersToLocal rejected: " + (Short $_.Exception.Message)) }

    # T07 enumerate everything
    try {
        Add-Line 'T07' 'INFO' ("{0} names in workbook collection" -f $wb.Names.Count)
        foreach ($x in $wb.Names) {
            $scope = try { $x.Parent.Name } catch { '?' }
            $cm = try { $x.Comment } catch { '' }
            Add-Line 'T07' 'ITEM' ("{0} | visible={1} | scope={2} | comment={3} | {4}" -f $x.Name, $x.Visible, $scope, $cm, (Short $x.RefersTo))
        }
    } catch { Add-Line 'T07' 'FAIL' $_.Exception.Message }

    # T19 save a copy; T20 is the file locked while open?
    try {
        $wb.SaveAs($saved, 51)
        Add-Line 'T19' 'PASS' "saved $saved"
        try { $fs = [System.IO.File]::Open($saved, 'Open', 'ReadWrite', 'None'); $fs.Close(); Add-Line 'T20' 'INFO' 'open workbook file is NOT locked' }
        catch { Add-Line 'T20' 'INFO' ('open workbook file is locked: ' + $_.Exception.Message) }
    } catch { Add-Line 'T19' 'FAIL' $_.Exception.Message }
    $wb.Close($false)
} finally {
    try { $xl.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}
}

# Stored form: <definedNames> in xl/workbook.xml
Add-Line 'X00' 'INFO' 'Stored form (xl/workbook.xml):'
if (Test-Path $saved) {
    try {
        $zip = Join-Path $stage 'probe_win.zip'; $ex = Join-Path $stage 'x'
        Copy-Item $saved $zip -Force; if (Test-Path $ex) { Remove-Item $ex -Recurse -Force }
        Expand-Archive $zip -DestinationPath $ex -Force
        $xml = Get-Content -Raw -Encoding UTF8 (Join-Path $ex 'xl\workbook.xml')
        foreach ($m in [regex]::Matches($xml, '<definedName ([^>]*)>(.*?)</definedName>', 'Singleline')) {
            $body = $m.Groups[2].Value -replace '&quot;', '"' -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>' -replace "`r", '\r' -replace "`n", '\n'
            Add-Line 'X01' 'ITEM' ($m.Groups[1].Value + ' => ' + (Short $body 160))
        }
        Copy-Item $saved $outDir -Force
    } catch { Add-Line 'X01' 'FAIL' $_.Exception.Message }
}

# ===== F: file-level route. Patch the copy saved above AS A ZIP (no Excel), then let Excel
# open it once. Same checks as probes/filelevel/probe_filelevel.py on the Mac:
#   F1 P_Add1 redefined as x+100 -> A1 must show 141 (cached 42)   F4 fullCalcOnLoad="1"
#   F2 Z_Add2 with a comment attribute   F3 Z_Loc scoped to S2 by localSheetId
#   F5 custom XML part with module source   F6 Z_Bare without its _xlfn. prefix
$patched = Join-Path $stage 'probe_patched_win.xlsx'
$resaved = Join-Path $stage 'probe_resaved_win.xlsx'
$utf8 = New-Object System.Text.UTF8Encoding($false)
function Read-Entry($e) { $r = New-Object System.IO.StreamReader($e.Open(), $utf8); try { $r.ReadToEnd() } finally { $r.Close() } }
function Write-Entry($z, $name, $text) { $w = New-Object System.IO.StreamWriter($z.CreateEntry($name).Open(), $utf8); try { $w.Write($text) } finally { $w.Close() } }

Add-Line 'F00' 'INFO' 'File-level route: patch probe_win.xlsx as a zip, then open it in Excel'
if (Test-Path $saved) {
    try {
        Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
        foreach ($p in $patched, $resaved) { if (Test-Path $p) { Remove-Item $p -Force } }
        $module = "// module: Probe`n/** Adds two. */`nZ_Add2 = LAMBDA(x, x + 2);`n"
        $zin = [System.IO.Compression.ZipFile]::OpenRead($saved)
        $zout = [System.IO.Compression.ZipFile]::Open($patched, 'Create')
        try {
            foreach ($e in $zin.Entries) {
                $n = $e.FullName
                if ($n -eq 'xl/workbook.xml') {
                    $x = Read-Entry $e
                    $x = [regex]::Replace($x, '(<definedName name="P_Add1"[^>]*>)[^<]*(</definedName>)', '${1}_xlfn.LAMBDA(_xlpm.x, _xlpm.x+100)${2}')
                    $sheets = @([regex]::Matches($x, '<sheet [^>]*name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
                    $s2i = [array]::IndexOf($sheets, 'S2')
                    $new = '<definedName name="Z_Add2" comment="doc written by file patch">_xlfn.LAMBDA(_xlpm.x, _xlpm.x+2)</definedName>' +
                           ('<definedName name="Z_Loc" localSheetId="{0}">7</definedName>' -f $s2i) +
                           '<definedName name="Z_Bare">SEQUENCE(1,3)</definedName>'
                    $x = $x.Replace('</definedNames>', $new + '</definedNames>')
                    $x = $x -replace '\s*fullCalcOnLoad="[^"]*"', ''
                    $x = $x -replace '<calcPr ', '<calcPr fullCalcOnLoad="1" '
                    Write-Entry $zout $n $x
                } elseif ($n -eq 'xl/_rels/workbook.xml.rels') {
                    Write-Entry $zout $n ((Read-Entry $e).Replace('</Relationships>', '<Relationship Id="rIdXln1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>'))
                } elseif ($n -eq '[Content_Types].xml') {
                    Write-Entry $zout $n ((Read-Entry $e).Replace('</Types>', '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>'))
                } else {
                    $src = $e.Open(); $dst = $zout.CreateEntry($n).Open()
                    try { $src.CopyTo($dst) } finally { $dst.Close(); $src.Close() }
                }
            }
            Write-Entry $zout 'customXml/item1.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xlnModule xmlns="urn:excel-dim:module:v0" name="Probe"><![CDATA[' + $module + ']]></xlnModule>')
            Write-Entry $zout 'customXml/itemProps1.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="{6C1A2B3C-0000-4000-8000-0000000000D1}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="urn:excel-dim:module:v0"/></ds:schemaRefs></ds:datastoreItem>')
            Write-Entry $zout 'customXml/_rels/item1.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>')
        } finally { $zout.Dispose(); $zin.Dispose() }
        Add-Line 'F00' 'PASS' ("patched copy written; S2 is sheet position {0}" -f $s2i)
    } catch { Add-Line 'F00' 'FAIL' ("patch: " + $_.Exception.Message) }
}

if (Test-Path $patched) {
    $xl2 = $null
    try { $xl2 = New-Object -ComObject Excel.Application } catch { Add-Line 'F01' 'FAIL' ("second Excel instance: " + $_.Exception.Message) }
    if ($xl2) {
        try {
            $xl2.Visible = $true; $xl2.DisplayAlerts = $false
            $wb2 = $xl2.Workbooks.Open($patched)
            Add-Line 'F01' 'INFO' ("opened; window caption '" + $xl2.ActiveWindow.Caption + "' (a repair would add [Riparato]/[Repaired])")
            $locks = @(Get-ChildItem $stage -Force -Filter '~$*' | ForEach-Object { $_.Name })
            Add-Line 'F01' ($(if ($locks.Count) { 'PASS' } else { 'FAIL' })) ("lock file while open: " + ($(if ($locks.Count) { $locks -join ', ' } else { 'NONE' })))
            $sh = $wb2.Worksheets.Item('S1')
            $a1 = $sh.Range('A1').Value2
            Add-Line 'F02' ($(if ($a1 -eq 141) { 'PASS' } else { 'FAIL' })) ("F1+F4 A1 =P_Add1(41) on open -> {0} (cached 42; 141 = recalculated on load)" -f $sh.Range('A1').Text)
            $cells = foreach ($a in 'B1', 'C2', 'A4', 'A5', 'A6', 'A7', 'A9', 'A10') { "$a=" + $sh.Range($a).Text }
            Add-Line 'F02' 'INFO' ("other cells: " + ($cells -join '  '))
            try { $v = $xl2.Evaluate('Z_Add2(1)'); $c = $wb2.Names.Item('Z_Add2').Comment
                  Add-Line 'F03' ($(if ($v -eq 3 -and $c -eq 'doc written by file patch') { 'PASS' } else { 'FAIL' })) ("F2 Z_Add2(1) -> {0}; comment '{1}'" -f $v, $c) } catch { Add-Line 'F03' 'FAIL' $_.Exception.Message }
            try { $v = $xl2.Evaluate('S2!Z_Loc'); $l = @(); foreach ($q in $wb2.Worksheets.Item('S2').Names) { $l += $q.Name }
                  Add-Line 'F04' ($(if ($v -eq 7) { 'PASS' } else { 'FAIL' })) ("F3 S2!Z_Loc -> {0}; names of S2: {1}" -f $v, ($l -join '; ')) } catch { Add-Line 'F04' 'FAIL' $_.Exception.Message }
            try { $sh.Range('H1').Formula2 = '=COLUMNS(Z_Bare)'; $xl2.Calculate()
                  Add-Line 'F05' 'INFO' ("F6 =COLUMNS(Z_Bare) -> {0} (expect #NAME?); RefersTo={1}" -f $sh.Range('H1').Text, $wb2.Names.Item('Z_Bare').RefersTo) } catch { Add-Line 'F05' 'FAIL' $_.Exception.Message }
            $wb2.SaveAs($resaved, 51); $wb2.Close($false)
        } catch { Add-Line 'F01' 'FAIL' $_.Exception.Message
        } finally {
            try { $xl2.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl2) } catch {}
        }
    }
}

if (Test-Path $resaved) {
    try {
        $z = [System.IO.Compression.ZipFile]::OpenRead($resaved)
        try {
            $x = Read-Entry ($z.GetEntry('xl/workbook.xml'))
            Add-Line 'F06' 'INFO' ("after Excel re-saved it: " + ([regex]::Match($x, '<calcPr[^>]*>').Value))
            foreach ($m in [regex]::Matches($x, '<definedName ([^>]*)>([^<]*)</definedName>')) {
                if ($m.Groups[1].Value -match 'name="(Z_|P_Add1)') { Add-Line 'F06' 'ITEM' ($m.Groups[1].Value + ' => ' + $m.Groups[2].Value) }
            }
            $cx = @($z.Entries | Where-Object { $_.FullName -like 'customXml/*' })
            $src = $false
            foreach ($e in $cx) { if ($e.FullName -match '^customXml/item\d+\.xml$' -and (Read-Entry $e) -match 'Z_Add2 = LAMBDA') { $src = $true } }
            Add-Line 'F07' ($(if ($src) { 'PASS' } else { 'FAIL' })) ("F5 customXml parts after re-save: " + ($(if ($cx.Count) { ($cx | ForEach-Object { $_.FullName }) -join ', ' } else { 'NONE (dropped)' })) + "; module source kept: $src")
        } finally { $z.Dispose() }
    } catch { Add-Line 'F06' 'FAIL' $_.Exception.Message }
}

$script:rpt | Set-Content -Encoding UTF8 $outFile
Write-Host ""; Write-Host "report: $outFile"
