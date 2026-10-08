#!/usr/bin/env python3
"""
probe_filelevel.py — P0 for the CODE-ONLY route: patch names into a saved .xlsx without
Excel, then let Excel open it once and see what survived.

    python3 probes/filelevel/probe_filelevel.py          (macOS: drives Excel to verify)

Base workbook: probes/results/probe_mac.xlsx (saved by Excel in the AppleScript probe;
S1!A1 holds =P_Add1(41) with cached value 42). The patch, done on the zip only:

  F1  redefine P_Add1 as x+100               -> after open, A1 should be 141 (recalc) not 42
  F2  add Z_Add2 = LAMBDA(x, x+2) with a comment="..." attribute
  F3  add a sheet-scoped name Z_Loc (localSheetId of S2), qualified as Excel stores it
  F4  set <calcPr fullCalcOnLoad="1"/>
  F5  add a custom XML part carrying the module source text (AFE-style "source in the file")
  F6  add Z_Bare = SEQUENCE(1,3) WITHOUT the _xlfn. prefix (what happens if the compiler forgets?)

Then Excel opens the patched copy (from the Office container: no file-access dialog),
we look for the ~$ lock file, read values, save-as a second copy, close, and diff what
Excel kept in xl/workbook.xml and customXml/.
"""
import pathlib, re, shutil, subprocess, sys, time, zipfile, html

HERE = pathlib.Path(__file__).resolve().parent
BASE = HERE.parent / "results" / "probe_mac.xlsx"
STAGE = pathlib.Path.home() / "Library/Group Containers/UBF8T346G9.Office/excel-dim-probe"
PATCHED = STAGE / "probe_patched.xlsx"
RESAVED = STAGE / "probe_resaved.xlsx"
OUT = HERE.parent / "results"

MODULE_SRC = """// module: Probe
/** Adds two. */
Z_Add2 = LAMBDA(x, x + 2);
"""

def patch(src: pathlib.Path, dst: pathlib.Path):
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}
    wb = files["xl/workbook.xml"].decode("utf-8")

    # F1: redefine P_Add1
    wb, n = re.subn(r'(<definedName name="P_Add1"[^>]*>)[^<]*(</definedName>)',
                    r'\g<1>_xlfn.LAMBDA(_xlpm.x, _xlpm.x+100)\g<2>', wb)
    assert n == 1, "P_Add1 not found"
    # sheet index of S2 in <sheets> order (localSheetId is the 0-based POSITION)
    sheets = re.findall(r'<sheet [^>]*name="([^"]+)"', wb)
    s2 = sheets.index("S2")
    new = (f'<definedName name="Z_Add2" comment="doc written by file patch">'
           f'_xlfn.LAMBDA(_xlpm.x, _xlpm.x+2)</definedName>'                       # F2
           f'<definedName name="Z_Loc" localSheetId="{s2}">7</definedName>'          # F3
           f'<definedName name="Z_Bare">SEQUENCE(1,3)</definedName>')                # F6
    wb = wb.replace("</definedNames>", new + "</definedNames>")
    # F4
    if "<calcPr" in wb:
        wb = re.sub(r"<calcPr([^/]*)/>", lambda m: "<calcPr" + re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wb)
    else:
        wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    files["xl/workbook.xml"] = wb.encode("utf-8")

    # F5: custom XML part, related from the workbook part
    item = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<xlnModule xmlns="urn:excel-dim:module:v0" name="Probe"><![CDATA['
            + MODULE_SRC + ']]></xlnModule>')
    props = ('<?xml version="1.0" encoding="UTF-8" standalone="no"?>'
             '<ds:datastoreItem ds:itemID="{6C1A2B3C-0000-4000-8000-0000000000D1}" '
             'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">'
             '<ds:schemaRefs><ds:schemaRef ds:uri="urn:excel-dim:module:v0"/></ds:schemaRefs></ds:datastoreItem>')
    files["customXml/item1.xml"] = item.encode("utf-8")
    files["customXml/itemProps1.xml"] = props.encode("utf-8")
    files["customXml/_rels/item1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/>'
        '</Relationships>').encode("utf-8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rels = rels.replace("</Relationships>",
        '<Relationship Id="rIdXln1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>')
    files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    ct = files["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace("</Types>",
        '<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>')
    files["[Content_Types].xml"] = ct.encode("utf-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():                      # keep the original part order
            zout.writestr(n, files.pop(n))
        for n, b in files.items():
            zout.writestr(n, b)

def osa(script: str, timeout=120) -> str:
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=timeout)
    return (r.stdout + r.stderr).strip()

def names_in(path: pathlib.Path):
    x = zipfile.ZipFile(path).read("xl/workbook.xml").decode("utf-8")
    calc = re.search(r"<calcPr[^>]*>", x)
    out = [f"calcPr: {calc.group(0) if calc else None}"]
    for m in re.finditer(r"<definedName ([^>]*)>(.*?)</definedName>", x, re.S):
        if m.group(1).startswith('name="Z_') or 'P_Add1' in m.group(1):
            out.append(f"{m.group(1)} => {html.unescape(m.group(2))}")
    return out

def main():
    STAGE.mkdir(parents=True, exist_ok=True)
    for p in (PATCHED, RESAVED):
        p.unlink(missing_ok=True)
    patch(BASE, PATCHED)
    rep = ["# excel-dim P0 — file-level route (macOS)", f"# {time.strftime('%FT%T')}", ""]
    was_running = subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True).returncode == 0

    rep.append("## Opened in Excel")
    rep.append("open: " + osa(f'tell application "Microsoft Excel" to open workbook workbook file name "{PATCHED}"', 90))
    time.sleep(2)
    locks = [p.name for p in STAGE.iterdir() if p.name.startswith("~$")]
    rep.append(f"L1 lock file while open: {locks or 'NONE'}")
    rd = osa('''tell application "Microsoft Excel"
  set wb to workbook "probe_patched.xlsx"
  set s1 to worksheet "S1" of wb
  set r to "F1 A1 (cached 42; 141 = recalculated on load) -> " & ((value of range "A1" of s1) as text) & linefeed
  try
    set r to r & "F2 evaluate Z_Add2(1) -> " & ((evaluate name "Z_Add2(1)") as text) & linefeed
  on error e
    set r to r & "F2 evaluate Z_Add2(1) FAILED " & e & linefeed
  end try
  try
    set r to r & "F3 evaluate S2!Z_Loc -> " & ((evaluate name "S2!Z_Loc") as text) & linefeed
  on error e
    set r to r & "F3 FAILED " & e & linefeed
  end try
  set formula2 of range "H1" of s1 to "=COLUMNS(Z_Bare)"
  calculate
  set r to r & "F6 =COLUMNS(Z_Bare) -> " & (string value of range "H1" of s1) & linefeed
  set r to r & "F6 Z_Bare references " & (references of named item "Z_Bare" of wb) & linefeed
  save workbook as wb filename "''' + str(RESAVED) + '''"
  close workbook "probe_resaved.xlsx" saving no
  return r
end tell''', 120)
    rep += rd.splitlines()
    if not was_running:
        osa('tell application "Microsoft Excel" to quit')

    rep += ["", "## workbook.xml as WE wrote it"] + names_in(PATCHED)
    if RESAVED.exists():
        rep += ["", "## workbook.xml after Excel re-saved it"] + names_in(RESAVED)
        z = zipfile.ZipFile(RESAVED)
        cx = [n for n in z.namelist() if n.startswith("customXml/")]
        rep.append(f"F5 customXml parts after re-save: {cx or 'NONE (dropped)'}")
        for n in cx:
            if re.match(r"customXml/item\d+\.xml$", n):
                body = z.read(n).decode("utf-8", "replace")
                rep.append(f"F5 {n} contains module source: {'Z_Add2 = LAMBDA' in body}")
        shutil.copy2(RESAVED, OUT / RESAVED.name)
    else:
        rep.append("re-save did not happen")
    shutil.copy2(PATCHED, OUT / PATCHED.name)
    text = "\n".join(rep)
    (OUT / f"filelevel-mac-{time.strftime('%Y%m%d-%H%M%S')}.txt").write_text(text, encoding="utf-8")
    print(text)

if __name__ == "__main__":
    main()
