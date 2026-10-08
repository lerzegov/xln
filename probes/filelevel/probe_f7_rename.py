#!/usr/bin/env python3
"""
probe_f7_rename.py — F7: can a name be renamed at FILE level, rewriting its token in
existing cell formulas, without Excel?

    python3 probes/filelevel/probe_f7_rename.py          (macOS: drives Excel as the oracle)

1. Excel builds f7_base.xlsx: names and formulas laid out as traps (below). Conditional
   formatting and data validation formulas that use the name are injected into the file.
2. ORACLE: Excel opens a copy and renames the names live in the Name Manager (T09 showed
   it rewrites dependents), then saves f7_oracle.xlsx.
3. PATCH: this script renames the same names in f7_base.xlsx as a zip, token by token, in
   <definedNames>, cell <f> elements, CF <formula> and DV <formula1>, and sets
   fullCalcOnLoad -> f7_patched.xlsx. calcChain.xml and every other part are untouched.
4. Excel opens f7_patched.xlsx (recalculates on load) and saves f7_patched_resaved.xlsx.
5. Compare: formulas of patched vs oracle, cell by cell; cached values of the re-saved
   patched file vs the base file (a rename must not change a single number).

Renames: Rate -> Growth (workbook), Fn -> Fx (workbook LAMBDA), S2!Loc -> Spot (sheet
scope; a workbook-level Loc with the same name must NOT change).
Traps: shared formulas, a spill anchor and a name over it, a prefix-sharing name (RateX),
the name inside a string, LET shadowing (stored _xlpm.Rate), INDIRECT("Rate"), lower case,
a sheet-qualified read, a name used in another name's definition, CF and DV formulas.
"""
import html, pathlib, re, shutil, subprocess, time, zipfile
from xml.sax.saxutils import escape

HERE = pathlib.Path(__file__).resolve().parent
STAGE = pathlib.Path.home() / "Library/Group Containers/UBF8T346G9.Office/excel-dim-probe"
OUT = HERE.parent / "results"
BASE0, BASE = STAGE / "f7_base0.xlsx", STAGE / "f7_base.xlsx"
ORACLE, PATCHED = STAGE / "f7_oracle.xlsx", STAGE / "f7_patched.xlsx"
RESAVED = STAGE / "f7_patched_resaved.xlsx"

# (old, new, scope) with scope None = workbook, else the sheet name
RENAMES = [("Rate", "Growth", None), ("Fn", "Fx", None), ("Loc", "Spot", "S2")]

BUILD = r'''
on run argv
  set out to item 1 of argv
  tell application "Microsoft Excel"
    set wb to make new workbook
    set s1 to worksheet 1 of wb
    set name of s1 to "S1"
    try
      set s2 to make new worksheet at end of wb
    on error
      set s2 to make new worksheet at wb
    end try
    set name of s2 to "S2"
    make new named item at wb with properties {name:"Rate", references:"=0.1"}
    make new named item at wb with properties {name:"RateX", references:"=0.5"}
    make new named item at wb with properties {name:"Rate2", references:"=Rate*2"}
    make new named item at wb with properties {name:"Fn", references:"=LAMBDA(x, x*Rate)"}
    make new named item at wb with properties {name:"Loc", references:"=3"}
    make new named item at wb with properties {name:"S2!Loc", references:"=7"}
    set value of range "A2:A11" of s1 to {{1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}}
    set formula2 of range "A1" of s1 to "=Rate*2"
    set formula of range "B2:B11" of s1 to "=A2*Rate"
    set formula2 of range "C1" of s1 to "=RateX+Rate"
    set formula2 of range "C2" of s1 to "=\"Rate is \"&Rate"
    set formula2 of range "C3" of s1 to "=LET(Rate, 5, Rate*2)"
    set formula2 of range "C4" of s1 to "=Fn(10)"
    set formula2 of range "C5" of s1 to "=S2!Loc+Loc"
    set formula2 of range "E1" of s1 to "=SEQUENCE(3)*Rate"
    make new named item at wb with properties {name:"Spl", references:"=S1!$E$1#"}
    set formula2 of range "C6" of s1 to "=SUM(Spl)"
    set formula2 of range "C7" of s1 to "=INDIRECT(\"Rate\")"
    set formula2 of range "C8" of s1 to "=rate*3"
    set formula2 of range "C9" of s1 to "=ROWS(E1#)*Rate"
    set formula2 of range "C10" of s1 to "=Rate2+Fn(1)"
    set formula2 of range "A1" of s2 to "=Loc"
    set formula2 of range "A2" of s2 to "=Rate+Loc"
    set formula of range "A3:A5" of s2 to "=Loc*ROW()"
    set formula2 of range "B1" of s2 to "=S1!A1+S2!Loc"
    calculate
    save workbook as wb filename out
    close workbook "f7_base0.xlsx" saving no
  end tell
  return "built"
end run
'''

ORACLE_AS = r'''
on run argv
  set src to item 1 of argv
  set out to item 2 of argv
  set r to ""
  tell application "Microsoft Excel"
    open workbook workbook file name src
    set wb to active workbook
    set n to count of named items of wb
    repeat with i from 1 to n
      set nm to name of named item i of wb
      set r to r & "before: " & nm & linefeed
    end repeat
    set name of named item "Rate" of wb to "Growth"
    set name of named item "Fn" of wb to "Fx"
    repeat with i from 1 to n
      set ni to named item i of wb
      if name of ni is "'S2'!Loc" or name of ni is "S2!Loc" then
        set name of ni to "Spot"
        exit repeat
      end if
    end repeat
    repeat with i from 1 to n
      set r to r & "after: " & (name of named item i of wb) & linefeed
    end repeat
    calculate
    save workbook as wb filename out
    close workbook "f7_oracle.xlsx" saving no
  end tell
  return r
end run
'''

RESAVE_AS = r'''
on run argv
  set src to item 1 of argv
  set out to item 2 of argv
  tell application "Microsoft Excel"
    open workbook workbook file name src
    set wb to active workbook
    set r to "opened as: " & (name of wb)
    save workbook as wb filename out
    close workbook "f7_patched_resaved.xlsx" saving no
  end tell
  return r
end run
'''


def osa(script: str, *args, timeout=180) -> str:
    r = subprocess.run(["osascript", "-", *map(str, args)], input=script,
                       capture_output=True, text=True, timeout=timeout)
    return (r.stdout + r.stderr).strip()


# ---- the rename: a small Excel formula scanner -------------------------------------
TOKEN = re.compile(r'''
    (?P<str>"(?:[^"]|"")*")                      # string literal
  | (?P<qsheet>'(?:[^']|'')*'!)                  # 'quoted sheet'!
  | (?P<struct>\[[^\]]*\](?:\])?)                # [..] structured ref / external book
  | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)      # number (so 1E5 is not an identifier)
  | (?P<err>\#[A-Za-z0-9/]+[!?]?)                # #REF!, #N/A, ... and the spill #
  | (?P<id>[A-Za-z_\\À-￿][\w.\\?À-￿]*!?)   # identifier, maybe Sheet!
  | (?P<other>.)
''', re.X | re.S)


def rename_formula(text: str, sheet: str | None, local: dict, renames) -> str:
    """Rename name tokens in one formula. `sheet` is the sheet the formula lives on (None
    for a workbook-scoped definedName), `local` maps sheet -> set of its local names."""
    out, qual, ext = [], None, False  # qual: the sheet qualifier just emitted, if any
    for m in TOKEN.finditer(text):
        kind, tok = m.lastgroup, m.group()
        prev = out[-1] if out else ""
        if kind == "qsheet":
            out.append(tok); qual = tok[1:-2].replace("''", "'")
            ext = qual.startswith("[") and not qual.startswith("[0]")
            continue
        if kind == "id" and tok.endswith("!"):
            out.append(tok); qual = tok[:-1]
            ext = prev.startswith("[") and prev != "[0]"   # [1]Sheet! = another workbook
            continue
        if kind == "other" and tok == "!" and prev.startswith("["):   # [1]!Name, [0]!Name
            out.append(tok); qual = prev[1:-1] == "0" and "" or prev
            ext = prev != "[0]"
            continue
        if kind == "id":
            for old, new, scope in renames:
                if tok.lower() != old.lower() or prev.startswith("[") or (qual is not None and ext):
                    continue
                ctx = (qual.removeprefix("[0]") or None) if qual is not None else sheet
                if scope is None:   # workbook name: not if a local one shadows it there
                    hit = not (ctx and old.lower() in local.get(ctx, set()))
                else:               # sheet name: only when the context is that sheet
                    hit = ctx == scope
                if hit:
                    tok = new
                    break
        out.append(tok); qual = None
    return "".join(out)


def patch_rename(src: pathlib.Path, dst: pathlib.Path) -> list[str]:
    log = []
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}
    wb = files["xl/workbook.xml"].decode("utf-8")
    sheets = re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb)
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    target = {rid: "xl/" + t.lstrip("/").removeprefix("xl/")
              for rid, t in re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
    target.update({rid: "xl/" + t.lstrip("/").removeprefix("xl/")
                   for t, rid in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    names = [html.unescape(s) for s, _ in sheets]
    local: dict[str, set] = {}
    for m in re.finditer(r'<definedName ([^>]*)>', wb):
        nm = re.search(r'name="([^"]+)"', m.group(1)).group(1)
        lid = re.search(r'localSheetId="(\d+)"', m.group(1))
        if lid:
            local.setdefault(names[int(lid.group(1))], set()).add(nm.lower())

    def fix(text, sheet):
        new = rename_formula(html.unescape(text), sheet, local, RENAMES)
        return escape(new)

    # definedNames: the name attribute, and the definition text
    def dn(m):
        attrs, body = m.group(1), m.group(2)
        lid = re.search(r'localSheetId="(\d+)"', attrs)
        sheet = names[int(lid.group(1))] if lid else None
        nm = re.search(r'name="([^"]+)"', attrs).group(1)
        for old, new, scope in RENAMES:
            if nm.lower() == old.lower() and scope == sheet:
                attrs = attrs.replace(f'name="{nm}"', f'name="{new}"')
                log.append(f"definedName {old} ({scope or 'workbook'}) -> {new}")
        return f"<definedName {attrs}>{fix(body, sheet)}</definedName>"
    wb = re.sub(r'<definedName ([^>]*)>(.*?)</definedName>', dn, wb, flags=re.S)
    wb = re.sub(r'\s*fullCalcOnLoad="[^"]*"', "", wb).replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    files["xl/workbook.xml"] = wb.encode("utf-8")

    for sname, rid in sheets:
        part = target[rid]
        x = files[part].decode("utf-8")
        before = x   # <f .../> (shared children) carry no text and must not open a match
        x = re.sub(r'(<f(?: [^>]*[^/])?>)(.*?)(</f>)', lambda m: m.group(1) + fix(m.group(2), sname) + m.group(3), x, flags=re.S)
        x = re.sub(r'(<formula>)(.*?)(</formula>)', lambda m: m.group(1) + fix(m.group(2), sname) + m.group(3), x, flags=re.S)
        x = re.sub(r'(<formula[12]>)(.*?)(</formula[12]>)', lambda m: m.group(1) + fix(m.group(2), sname) + m.group(3), x, flags=re.S)
        files[part] = x.encode("utf-8")
        log.append(f"{sname} ({part}): {'changed' if x != before else 'unchanged'}")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            zout.writestr(n, files[n])
    return log


# ---- inject CF + DV formulas into S1 (Excel's AppleScript cannot make them reliably) --
def inject_cf_dv(src: pathlib.Path, dst: pathlib.Path):
    zin = zipfile.ZipFile(src)
    files = {n: zin.read(n) for n in zin.namelist()}
    p = "xl/worksheets/sheet1.xml"
    x = files[p].decode("utf-8")
    cf = ('<conditionalFormatting sqref="A2:A11"><cfRule type="expression" priority="1">'
          '<formula>A2&gt;Rate*50</formula></cfRule></conditionalFormatting>'
          '<dataValidations count="1"><dataValidation type="decimal" operator="lessThan" '
          'allowBlank="1" showErrorMessage="1" sqref="G1"><formula1>Rate+Loc</formula1>'
          '</dataValidation></dataValidations>')
    assert "<pageMargins" in x
    files[p] = x.replace("<pageMargins", cf + "<pageMargins", 1).encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in zin.namelist():
            zout.writestr(n, files[n])


# ---- reading a workbook for comparison -------------------------------------------------
def snapshot(path: pathlib.Path):
    z = zipfile.ZipFile(path)
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    tgt = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    tgt.update({rid: t for t, rid in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    sheets = re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb)
    names = [s for s, _ in sheets]
    snap = {"names": {}, "f": {}, "v": {}, "cf": {}, "shared": 0, "calcChain": "xl/calcChain.xml" in z.namelist()}
    for m in re.finditer(r'<definedName ([^>]*)>(.*?)</definedName>', wb, re.S):
        nm = re.search(r'name="([^"]+)"', m.group(1)).group(1)
        lid = re.search(r'localSheetId="(\d+)"', m.group(1))
        key = (names[int(lid.group(1))] + "!" if lid else "") + nm
        if not nm.startswith("_xl"):
            snap["names"][key] = html.unescape(m.group(2))
    for sname, rid in sheets:
        x = z.read("xl/" + tgt[rid].lstrip("/").removeprefix("xl/")).decode("utf-8")
        for c in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', x, re.S):
            ref, inner = f"{sname}!{c.group(1)}", c.group(3) or ""
            f = re.search(r'<f([^>]*?)(?:/>|>(.*?)</f>)', inner, re.S)
            if f:
                if 't="shared"' in f.group(1):
                    snap["shared"] += 1
                snap["f"][ref] = html.unescape(f.group(2)) if f.group(2) is not None else "(shared child)"
            v = re.search(r'<v>(.*?)</v>', inner, re.S)
            if v:
                snap["v"][ref] = html.unescape(v.group(1))
        for i, m in enumerate(re.finditer(r'<(formula[12]?)>(.*?)</\1>', x, re.S)):
            snap["cf"][f"{sname} {m.group(1)}#{i}"] = html.unescape(m.group(2))
    return snap


def same_number(a, b):
    try:
        return abs(float(a) - float(b)) <= 1e-12 * max(1, abs(float(a)))
    except ValueError:
        return a == b


def main():
    STAGE.mkdir(parents=True, exist_ok=True)
    for p in (BASE0, BASE, ORACLE, PATCHED, RESAVED):
        p.unlink(missing_ok=True)
    was_running = subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True).returncode == 0
    rep = ["# excel-dim F7 — rename a name inside cell formulas at file level (macOS)",
           f"# {time.strftime('%FT%T')}", ""]

    rep.append("## 1 build: " + osa(BUILD, BASE0))
    inject_cf_dv(BASE0, BASE)
    rep.append("## 2 oracle (Excel renames live):")
    rep += ["   " + l for l in osa(ORACLE_AS, BASE, ORACLE).splitlines()]
    rep.append("## 3 file patch:")
    rep += ["   " + l for l in patch_rename(BASE, PATCHED)]
    rep.append("## 4 Excel opens the patched file: " + osa(RESAVE_AS, PATCHED, RESAVED))
    if not was_running:
        osa('tell application "Microsoft Excel" to quit')

    b, o, p = snapshot(BASE), snapshot(ORACLE), snapshot(PATCHED)
    r = snapshot(RESAVED) if RESAVED.exists() else None
    rep += ["", f"shared-formula cells in base: {b['shared']}; calcChain present: base {b['calcChain']}, "
            f"patched {p['calcChain']}, resaved {r and r['calcChain']}"]

    rep += ["", "## 5a formulas: patched vs oracle (Excel's own rename)"]
    diffs = 0
    for k in sorted(set(o["f"]) | set(p["f"])):
        of, pf = o["f"].get(k), p["f"].get(k)
        mark = "OK  " if of == pf else "DIFF"
        diffs += of != pf
        rep.append(f"{mark} {k:8} base={b['f'].get(k)!r:30} oracle={of!r:30} patched={pf!r}")
    rep += ["", "## 5b names: patched vs oracle"]
    for k in sorted(set(o["names"]) | set(p["names"])):
        of, pf = o["names"].get(k), p["names"].get(k)
        diffs += of != pf
        rep.append(f"{'OK  ' if of == pf else 'DIFF'} {k:10} oracle={of!r:28} patched={pf!r}")
    rep += ["", "## 5c CF / DV formulas: patched vs oracle"]
    for k in sorted(set(o["cf"]) | set(p["cf"])):
        of, pf = o["cf"].get(k), p["cf"].get(k)
        diffs += of != pf
        rep.append(f"{'OK  ' if of == pf else 'DIFF'} {k:18} base={b['cf'].get(k)!r:20} oracle={of!r:24} patched={pf!r}")
    rep.append(f"=> {diffs} difference(s) from Excel's own rename")

    if r:
        rep += ["", "## 5d values after Excel recalculated the patched file vs base (must not change)"]
        vd = 0
        for k in sorted(b["v"]):
            if k not in b["f"]:
                continue
            same = same_number(b["v"][k], r["v"].get(k, "<none>"))
            vd += not same
            rep.append(f"{'OK  ' if same else 'DIFF'} {k:8} base={b['v'][k]!r:22} after={r['v'].get(k)!r}")
        rep.append(f"=> {vd} changed value(s)")
        rep += ["", "## 5e formulas after Excel re-saved the patched file vs patched (does Excel keep them?)"]
        kd = 0
        for k in sorted(p["f"]):
            if r["f"].get(k) != p["f"][k]:
                kd += 1
                rep.append(f"DIFF {k:8} patched={p['f'][k]!r} resaved={r['f'].get(k)!r}")
        rep.append(f"=> {kd} formula(s) rewritten by Excel on save")

    for f in (BASE, ORACLE, PATCHED, RESAVED):
        if f.exists():
            shutil.copy2(f, OUT / f.name)
    text = "\n".join(rep)
    (OUT / f"f7-mac-{time.strftime('%Y%m%d-%H%M%S')}.txt").write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
