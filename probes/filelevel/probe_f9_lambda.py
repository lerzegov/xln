#!/usr/bin/env python3
"""
probe_f9_lambda.py — F9: how Excel stores LAMBDA forms, in defined names and in cells.

    /opt/homebrew/bin/python3 probes/filelevel/probe_f9_lambda.py      (macOS, Excel running)

Excel makes a new workbook, enters every form below live (named items with `references`,
cells with `formula2`), calculates and saves it as probes/results/f9_lambda_mac.xlsx. The
script then prints each name's stored definition and each cell's <f> and <v>, as Excel
wrote them; probes/README.md (F9) records them. A form Excel refuses is reported, not
retried. Only the workbook this script creates is touched.
"""
import pathlib, shutil, subprocess, sys, zipfile
from xml.dom import minidom

HERE = pathlib.Path(__file__).resolve().parent
STAGE = pathlib.Path.home() / "Library/Group Containers/UBF8T346G9.Office/xln-m3d"
OUT = HERE.parent / "results" / "f9_lambda_mac.xlsx"
FILE = "f9_lambda_mac.xlsx"

# (name, definition as typed in the Name Manager)
NAMES = [
    ("ANA.GROW", "=LAMBDA(value, rate, [periods], value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods))"),
    ("TWOOPT", "=LAMBDA(a, [b], [d], a + IF(ISOMITTED(b), 0, b) + IF(ISOMITTED(d), 0, d))"),
    ("ALLOPT", "=LAMBDA([x], [y], IF(ISOMITTED(x), 1, x) * IF(ISOMITTED(y), 2, y))"),
    ("LETIN", "=LAMBDA(x, LET(y, x * 2, z, y + 1, y * z))"),
    ("INLET", "=LET(f, LAMBDA(t, t * 3), f(2))"),
    ("IMMED", "=LAMBDA(x, x + 1)(2)"),
    ("FACT", "=LAMBDA(n, IF(n <= 1, 1, n * FACT(n - 1)))"),
    ("MAKEADDER", "=LAMBDA(n, LAMBDA(x, x + n))"),
    ("ETAMAP", "=MAP(Sheet1!$A$1:$A$3, ABS)"),
    ("ETABYROW", "=BYROW(Sheet1!$A$1:$A$3, SUM)"),
    ("REDUCED", "=REDUCE(0, Sheet1!$A$1:$A$3, LAMBDA(acc, v, acc + v))"),
    ("OPTNEST", "=LAMBDA(x, [k], LET(m, IF(ISOMITTED(k), 2, k), LAMBDA(y, [z], x * m + y + IF(ISOMITTED(z), 0, z))))"),
    ("ANA.CUBE", "=LAMBDA(x, x ^ 3)"),
]

# (cell, formula2)
CELLS = [
    ("B1", "=ANA.GROW(100, 0.1, 2)"),
    ("B2", "=ANA.GROW(100, 0.1)"),
    ("B3", "=TWOOPT(1)"),
    ("B4", "=TWOOPT(1, 2, 3)"),
    ("B5", "=ALLOPT()"),
    ("B6", "=ALLOPT(5)"),
    ("B7", "=ALLOPT(, 7)"),
    ("B8", "=LETIN(3)"),
    ("B9", "=INLET"),
    ("B10", "=IMMED"),
    ("B11", "=FACT(5)"),
    ("B12", "=MAKEADDER(2)(3)"),
    ("B13", "=SUM(ETAMAP)"),
    ("B14", "=SUM(ETABYROW)"),
    ("B15", "=REDUCED"),
    ("B16", "=OPTNEST(3)(1)"),
    ("B17", "=OPTNEST(3, 4)(1, 5)"),
    ("B18", "=ANA.CUBE(2)"),
    ("D1", "=LAMBDA(x, [y], x + IF(ISOMITTED(y), 10, y))(1)"),
    ("D2", "=LET(f, LAMBDA(t, t * 3), f(2))"),
    ("D3", "=LAMBDA(x, LET(y, x * 2, y + 1))(4)"),
    ("D4", "=REDUCE(0, A1:A3, LAMBDA(a, v, a + v))"),
    ("D5", "=LAMBDA(n, LAMBDA(x, x + n))(2)(3)"),
    ("D6", "=LAMBDA([p], IF(ISOMITTED(p), \"none\", p))()"),
    ("F1", "=MAP(A1:A3, ABS)"),
    ("G1", "=BYROW(A1:A3, SUM)"),
    ("H1", "=SCAN(0, A1:A3, LAMBDA(a, v, a + v))"),
]

BUILD = r'''
on run argv
  set out to item 1 of argv
  set report to ""
  tell application "Microsoft Excel"
    set wb to make new workbook
    set s1 to worksheet 1 of wb
    set name of s1 to "Sheet1"
    set value of range "A1:A3" of s1 to {{-1}, {2}, {-3}}
%NAMES%
%CELLS%
    calculate
    save workbook as wb filename out
    close workbook "%FILE%" saving no
  end tell
  return report
end run
'''


def as_str(s: str) -> str:
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def script() -> str:
    names = "\n".join(
        f'    try\n      make new named item at wb with properties {{name:{as_str(n)}, references:{as_str(d)}}}\n'
        f'    on error msg\n      set report to report & "name {n}: " & msg & linefeed\n    end try'
        for n, d in NAMES)
    cells = "\n".join(
        f'    try\n      set formula2 of range "{c}" of s1 to {as_str(f)}\n'
        f'    on error msg\n      set report to report & "cell {c}: " & msg & linefeed\n    end try'
        for c, f in CELLS)
    return BUILD.replace("%NAMES%", names).replace("%CELLS%", cells).replace("%FILE%", FILE)


def text(node) -> str:
    return "".join(c.data for c in node.childNodes if c.nodeType in (c.TEXT_NODE, c.CDATA_SECTION_NODE))


def main() -> int:
    STAGE.mkdir(parents=True, exist_ok=True)
    staged = STAGE / FILE
    if staged.exists():
        staged.unlink()
    try:
        r = subprocess.run(["osascript", "-", str(staged)], input=script(), capture_output=True, text=True, timeout=120)
    except subprocess.TimeoutExpired:
        print("TIMEOUT: Excel did not answer (a dialog is probably open)")
        return 1
    out = (r.stdout + r.stderr).strip()
    if out:
        print("Excel reported:\n" + out)
    if not staged.exists():
        print("no workbook saved")
        return 1
    shutil.copyfile(staged, OUT)
    with zipfile.ZipFile(OUT) as z:
        wb = minidom.parseString(z.read("xl/workbook.xml"))
        print("\n# names")
        for d in wb.getElementsByTagName("definedName"):
            print(f"{d.getAttribute('name')} = {text(d)!r}")
        sh = minidom.parseString(z.read("xl/worksheets/sheet1.xml"))
        print("\n# cells")
        for c in sh.getElementsByTagName("c"):
            fs = c.getElementsByTagName("f")
            if not fs:
                continue
            f = fs[0]
            vs = c.getElementsByTagName("v")
            attrs = " ".join(f'{k}="{v}"' for k, v in f.attributes.items())
            cm = c.getAttribute("cm")
            print(f"{c.getAttribute('r')}{' cm=' + cm if cm else ''} <f {attrs}> {text(f)!r} -> {text(vs[0]) if vs else None!r}")
    print(f"\nwrote {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
