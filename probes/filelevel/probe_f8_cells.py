#!/usr/bin/env python3
"""
probe_f8_cells.py — F8: can cell formulas be replaced at FILE level, without Excel, and
does Excel open the result without repair and recalculate it correctly?

    /opt/homebrew/bin/python3 probes/filelevel/probe_f8_cells.py [--with-repair]   (macOS)

--with-repair also opens p2, p2b, p2c (stale calcChain): each makes Excel show its repair
prompt, which someone must answer (Yes writes the repair log, which the script reads).

Repairs are detected from Excel's repair log ('Repair Result to <file><n>.xml' in its
sandbox tmp folder); a dialog that blocks the open is caught by a timeout.

1. Excel builds f8_base.xlsx (sheets N, D, Slot, Sh) and f8_plain_base.xlsx (no formulas,
   so no metadata.xml and no calcChain.xml).
2. ORACLE: Excel opens a copy of the base, makes every edit live (formula2), saves
   f8_oracle.xlsx. Its values are read the same way as the patched files'.
3. PATCH: this script edits the base as a zip with a real XML parser (xml.dom.minidom),
   in several variants that differ only in the workbook-wide choices of cases 6 and 7:
     p1  drop calcChain, no <v> on edited cells, fullCalcOnLoad="1"   (the candidate recipe)
     p2  keep the stale calcChain, otherwise p1
     p2b keep it minus the entries of cells that no longer hold a formula
     p2c also minus the entries whose array flag (a="1") no longer matches the cell
     p3  keep the stale <v> of edited cells, otherwise p1
     p4  no <v>, no fullCalcOnLoad
     p5  stale <v>, no fullCalcOnLoad
     p6  stale <v>, no fullCalcOnLoad, calcId set to the current Excel's (no recalc reason)
     p7  no <v>, no fullCalcOnLoad, current calcId
   and the plain workbook in two: q1 (dynamic-array form, metadata.xml created) and q2
   (t="array" without cm, no metadata.xml: a legacy CSE array).
4. Excel opens each patched file; the script reads formula2 and the displayed text of
   every watched cell before anything else happens, then saves a re-saved copy.
5. Compare: values on open vs the oracle; stored XML (f text, t, ref, si, cm, <v>) of the
   re-saved file vs the oracle; calcChain and metadata parts.

Cases (PLAN.md, probe F8):
  N     1 normal cell formulas; 4 value and shared-string cells turned into formulas;
        traps: an array-evaluating formula written without cm, a scalar written with cm.
  D     2 spill anchors: bigger and smaller spill, the old spill's ghost cells removed or
        kept, the old ref kept or reduced to the anchor; spill -> scalar; scalar -> spill.
        Names on 'D'!$C$6# and 'D'!$F$6#.
  Slot  3 empty named cells from Create from Selection, filled with a spilling formula and
        with scalar ones, one in a row that does not exist in the file.
  Sh    5 shared formulas: edit a child; edit a master (promote the next child; or un-share
        the group); edit the last child and shrink the master's ref.
"""
import pathlib, shutil, subprocess, sys, time, zipfile
from xml.dom import minidom

HERE = pathlib.Path(__file__).resolve().parent
STAGE = pathlib.Path.home() / "Library/Group Containers/UBF8T346G9.Office/excel-dim-probe"
OUT = HERE.parent / "results"
BASE, PLAIN = STAGE / "f8_base.xlsx", STAGE / "f8_plain_base.xlsx"
ORACLE = STAGE / "f8_oracle.xlsx"
# The calcId this Excel writes on save. A new workbook saved by AppleScript carries
# 181029, so p4/p5 alone cannot tell fullCalcOnLoad from Excel's recalc of an older calcId.
CURRENT_CALC_ID = "191029"
VARIANTS = {   # name: (calcChain: drop | keep | prune | prune+a, keep stale <v>, fullCalcOnLoad, calcId)
    "p1": ("drop", False, True, None),
    "p3": ("drop", True, True, None),
    "p4": ("drop", False, False, None),
    "p5": ("drop", True, False, None),
    "p6": ("drop", True, False, CURRENT_CALC_ID),
    "p7": ("drop", False, False, CURRENT_CALC_ID),
    "p2": ("keep", False, True, None),
    "p2b": ("prune", False, True, None),
    "p2c": ("prune+a", False, True, None),
}
PLAIN_VARIANTS = ("q1", "q2")
XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

BUILD = r'''
on run argv
  set out to item 1 of argv
  set out2 to item 2 of argv
  tell application "Microsoft Excel"
    set wb to make new workbook
    repeat 3 times
      make new worksheet at end of wb
    end repeat
    set name of worksheet 1 of wb to "N"
    set name of worksheet 2 of wb to "D"
    set name of worksheet 3 of wb to "Slot"
    set name of worksheet 4 of wb to "Sh"
    set sN to worksheet "N" of wb
    set sD to worksheet "D" of wb
    set sS to worksheet "Slot" of wb
    set sH to worksheet "Sh" of wb
    -- N: normal formulas, values and text
    set value of range "A1:A3" of sN to {{1}, {2}, {3}}
    set formula2 of range "B1" of sN to "=SUM(A1:A3)"
    set formula2 of range "B2" of sN to "=A1*10"
    set formula2 of range "B3" of sN to "=A3"
    set formula2 of range "B4" of sN to "=A1"
    set formula2 of range "B5" of sN to "=A1"
    set formula2 of range "B6" of sN to "=A2"
    set formula2 of range "B7" of sN to "=A2"
    set value of range "C1" of sN to 5
    set value of range "C2" of sN to "hello"
    set value of range "C3" of sN to 7
    set value of range "C4" of sN to "x"
    -- D: spill anchors and names on them
    set formula2 of range "C6" of sD to "=SEQUENCE(3)"
    set formula2 of range "F6" of sD to "=SEQUENCE(4)"
    set formula2 of range "H6" of sD to "=SEQUENCE(3)"
    set formula2 of range "J6" of sD to "=SEQUENCE(4)"
    set formula2 of range "L6" of sD to "=SEQUENCE(4)"
    set formula2 of range "N6" of sD to "=SEQUENCE(3)"
    set formula2 of range "P6" of sD to "=3+4"
    make new named item at wb with properties {name:"Spill", references:"='D'!$C$6#"}
    make new named item at wb with properties {name:"Spill2", references:"='D'!$F$6#"}
    set formula2 of range "E1" of sD to "=SUM(Spill)"
    set formula2 of range "E2" of sD to "=SUM(Spill2)"
    set formula2 of range "E3" of sD to "=ROWS(F6#)"
    set formula2 of range "E4" of sD to "=ROWS(C6#)"
    set formula2 of range "E5" of sD to "=SUM(H6#)"
    set formula2 of range "E6" of sD to "=SUM(J6#)"
    set formula2 of range "E7" of sD to "=SUM(L6#)"
    -- Slot: labels and names on empty cells, as Create from Selection makes them
    set value of range "A1" of sS to "Revenue"
    set value of range "A2" of sS to "Costs"
    set value of range "A6" of sS to "Tax"
    create names range "A1:B2" of sS left position true
    create names range "A6:B6" of sS left position true
    make new named item at wb with properties {name:"Far", references:"=Slot!$B$9"}
    -- Sh: shared formula groups
    set value of range "A1:A5" of sH to {{1}, {2}, {3}, {4}, {5}}
    set formula of range "B1:B5" of sH to "=A1*2"
    set formula of range "C1:C5" of sH to "=A1+1"
    set formula of range "D1:D5" of sH to "=A1*3"
    set formula of range "E1:E5" of sH to "=A1*4"
    calculate
    save workbook as wb filename out
    close workbook "f8_base.xlsx" saving no

    set wb to make new workbook
    set sP to worksheet 1 of wb
    set name of sP to "P"
    set value of range "A1:A3" of sP to {{1}, {2}, {3}}
    make new named item at wb with properties {name:"Out", references:"=P!$B$1"}
    save workbook as wb filename out2
    close workbook "f8_plain_base.xlsx" saving no
  end tell
  return "built"
end run
'''

ORACLE_AS = r'''
on run argv
  set src to item 1 of argv
  set out to item 2 of argv
  tell application "Microsoft Excel"
    open workbook workbook file name src
    set wb to active workbook
    set sN to worksheet "N" of wb
    set sD to worksheet "D" of wb
    set sS to worksheet "Slot" of wb
    set sH to worksheet "Sh" of wb
    set formula2 of range "B1" of sN to "=SUM(A1:A3)*100"
    set formula2 of range "B2" of sN to "=LET(x,A1*10,x+1)"
    clear contents range "B3" of sN
    set formula2 of range "B4" of sN to "=SUM(A1:A3*2)"
    set formula2 of range "B5" of sN to "=SUM(A1:A3*2)"
    set formula2 of range "B6" of sN to "=1+2"
    set formula2 of range "B7" of sN to "=LAMBDA(x,x*2)(A3)"
    set formula2 of range "C1" of sN to "=A1+100"
    set formula2 of range "C2" of sN to "=UPPER(\"ab\")&A1"
    set formula2 of range "C3" of sN to "=XLOOKUP(2,A1:A3,A1:A3)*7"
    set formula2 of range "C4" of sN to "=A1:A3*3"
    set formula2 of range "C6" of sD to "=SEQUENCE(5)"
    set formula2 of range "F6" of sD to "=SEQUENCE(2)"
    set formula2 of range "H6" of sD to "=SEQUENCE(5)"
    set formula2 of range "J6" of sD to "=SEQUENCE(2)"
    set formula2 of range "L6" of sD to "=SEQUENCE(2)"
    set formula2 of range "N6" of sD to "=42"
    set formula2 of range "P6" of sD to "=SEQUENCE(3)"
    set formula2 of range "B1" of sS to "=SEQUENCE(1,3)*10"
    set references of named item "Revenue" of wb to "='Slot'!$B$1#"
    set formula2 of range "B2" of sS to "=SUM(B1#)/2"
    set formula2 of range "B6" of sS to "=Costs*0.25"
    set formula2 of range "B9" of sS to "=Tax+1"
    set formula2 of range "F1" of sS to "=SUM(Revenue)"
    set formula2 of range "B3" of sH to "=A3*100"
    set formula2 of range "C1" of sH to "=A1+1000"
    set formula2 of range "D1" of sH to "=A1+2000"
    set formula2 of range "E5" of sH to "=A5*400"
    calculate
    save workbook as wb filename out
    close workbook "f8_oracle.xlsx" saving no
  end tell
  return "oracle saved"
end run
'''

PLAIN_ORACLE_AS = r'''
on run argv
  set src to item 1 of argv
  set out to item 2 of argv
  tell application "Microsoft Excel"
    open workbook workbook file name src
    set wb to active workbook
    set formula2 of range "B1" of worksheet "P" of wb to "=SEQUENCE(3)*A1"
    set formula2 of range "C1" of worksheet "P" of wb to "=SUM(B1#)"
    calculate
    save workbook as wb filename out
    close workbook "f8_plain_oracle.xlsx" saving no
  end tell
  return "plain oracle saved"
end run
'''

# Open, report what the user would see (no calculate first), optionally re-save, close.
READ_AS = r'''
on run argv
  set src to item 1 of argv
  set out to item 2 of argv
  set cells to item 3 of argv
  set AppleScript's text item delimiters to ","
  set refs to text items of cells
  set AppleScript's text item delimiters to ""
  set want to item 4 of argv
  set TB to tab -- inside the Excel block, `tab` is not the constant
  tell application "Microsoft Excel"
    with timeout of 60 seconds
      open workbook workbook file name src
    end timeout
    -- a repair prompt leaves no workbook of that name (or blocks the open: timeout)
    set wb to missing value
    repeat with i from 1 to (count of workbooks)
      if (name of workbook i) is want then set wb to workbook want
    end repeat
    if wb is missing value then return "NOT OPENED: no workbook named " & want
    set nm to name of wb
    set r to "workbook " & nm & ", sheets " & (count of worksheets of wb) & linefeed
    repeat with x in refs
      set AppleScript's text item delimiters to "!"
      set parts to text items of (x as text)
      set AppleScript's text item delimiters to ""
      set rg to range (item 2 of parts) of worksheet (item 1 of parts) of wb
      set f to formula2 of rg
      set s to string value of rg
      set r to r & "@" & (x as text) & TB & f & TB & s & linefeed
    end repeat
    if out is not "-" then
      save workbook as wb filename out
      set AppleScript's text item delimiters to "/"
      set nm to last text item of out
      set AppleScript's text item delimiters to ""
    end if
    close workbook nm saving no
  end tell
  return r
end run
'''

CELLS = (["N!" + c for c in "A1 A2 A3 B1 B2 B3 B4 B5 B6 B7 C1 C2 C3 C4 C5 C6".split()]
         + ["D!E%d" % i for i in range(1, 8)]
         + ["D!%s%d" % (col, r) for col in "CFHJLNP" for r in range(6, 11)]
         + ["Slot!" + c for c in "B1 C1 D1 B2 B6 B9 F1".split()]
         + ["Sh!%s%d" % (col, r) for col in "BCDE" for r in range(1, 6)])
PLAIN_CELLS = ["P!B1", "P!B2", "P!B3", "P!C1"]


def osa(script: str, *args, timeout=150) -> str:
    try:
        r = subprocess.run(["osascript", "-", *map(str, args)], input=script,
                           capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        return "TIMEOUT: Excel did not answer (a repair or alert dialog is probably open)"
    return (r.stdout + r.stderr).strip()


# ---- cell addresses ---------------------------------------------------------------
def split_ref(ref: str):
    i = 0
    while ref[i].isalpha():
        i += 1
    col = 0
    for ch in ref[:i].upper():
        col = col * 26 + ord(ch) - 64
    return col, int(ref[i:])


def col_name(col: int) -> str:
    s = ""
    while col:
        col, rem = divmod(col - 1, 26)
        s = chr(65 + rem) + s
    return s


def cells_in(rng: str):
    a, _, b = rng.partition(":")
    (c1, r1), (c2, r2) = split_ref(a), split_ref(b or a)
    return [col_name(c) + str(r) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


# ---- the package: parts as DOMs, written back with Excel's declaration -------------
class Package:
    def __init__(self, path: pathlib.Path):
        z = zipfile.ZipFile(path)
        self.order = z.namelist()
        self.raw = {n: z.read(n) for n in self.order}
        self.dom = {}

    def xml(self, part):
        if part not in self.dom:
            self.dom[part] = minidom.parseString(self.raw[part])
        return self.dom[part]

    def remove(self, part):
        self.order.remove(part)
        self.raw.pop(part)
        self.dom.pop(part, None)

    def add(self, part, text: str):
        self.order.append(part)
        self.raw[part] = text.encode("utf-8")

    def save(self, path: pathlib.Path):
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                if n in self.dom:
                    data = (XML_DECL + self.dom[n].documentElement.toxml()).encode("utf-8")
                else:
                    data = self.raw[n]
                z.writestr(n, data)

    def sheet_parts(self):
        wb = self.xml("xl/workbook.xml")
        rels = self.xml("xl/_rels/workbook.xml.rels")
        target = {r.getAttribute("Id"): r.getAttribute("Target")
                  for r in rels.getElementsByTagName("Relationship")}
        out = {}
        for s in wb.getElementsByTagName("sheet"):
            t = target[s.getAttribute("r:id")]
            out[s.getAttribute("name")] = t.lstrip("/") if t.startswith("/") else "xl/" + t
        return out


def children(node, tag):
    return [c for c in node.childNodes if c.nodeType == c.ELEMENT_NODE and c.tagName == tag]


def child(node, tag):
    found = children(node, tag)
    return found[0] if found else None


def text_of(node):
    return "".join(t.data for t in node.childNodes if t.nodeType == t.TEXT_NODE)


class Sheet:
    def __init__(self, doc):
        self.doc = doc
        self.data = doc.getElementsByTagName("sheetData")[0]

    def find(self, ref):
        col, row = split_ref(ref)
        for r in children(self.data, "row"):
            if int(r.getAttribute("r")) == row:
                for c in children(r, "c"):
                    if c.getAttribute("r") == ref:
                        return c
        return None

    def cell(self, ref):
        """The <c> element, created in row and column order when absent."""
        col, row = split_ref(ref)
        rows = children(self.data, "row")
        rnode = next((r for r in rows if int(r.getAttribute("r")) == row), None)
        if rnode is None:
            rnode = self.doc.createElement("row")
            rnode.setAttribute("r", str(row))
            after = next((r for r in rows if int(r.getAttribute("r")) > row), None)
            self.data.insertBefore(rnode, after)
        cs = children(rnode, "c")
        c = next((c for c in cs if c.getAttribute("r") == ref), None)
        if c is None:
            c = self.doc.createElement("c")
            c.setAttribute("r", ref)
            after = next((x for x in cs if split_ref(x.getAttribute("r"))[0] > col), None)
            rnode.insertBefore(c, after)
        return c

    def delete(self, ref):
        c = self.find(ref)
        if c is not None:
            c.parentNode.removeChild(c)

    def set_formula(self, ref, text, array=False, keep_v=False, shared=None, sst=None):
        """Replace whatever the cell holds by a formula. `array`: dynamic-array form
        (t="array" ref=<anchor> plus cm="1"); "cse": t="array" without cm. `shared`:
        (si, ref) to make the cell a shared-formula master."""
        c = self.cell(ref)
        old_v = child(c, "v")
        old_t = c.getAttribute("t")
        for tag in ("f", "v", "is"):
            for n in children(c, tag):
                c.removeChild(n)
        if c.hasAttribute("t"):
            c.removeAttribute("t")
        if c.hasAttribute("cm"):
            c.removeAttribute("cm")
        f = self.doc.createElement("f")
        if array:
            f.setAttribute("t", "array")
            f.setAttribute("ref", ref)
            if array is True:
                c.setAttribute("cm", "1")
        if shared:
            f.setAttribute("t", "shared")
            f.setAttribute("ref", shared[1])
            f.setAttribute("si", shared[0])
        f.appendChild(self.doc.createTextNode(text))
        c.appendChild(f)
        if keep_v and old_v is not None:
            v = old_v
            if old_t == "s":   # a shared-string cell keeps its old text as a string result
                c.setAttribute("t", "str")
                v = self.doc.createElement("v")
                v.appendChild(self.doc.createTextNode(sst[int(text_of(old_v))]))
            elif old_t in ("str", "e", "b"):
                c.setAttribute("t", old_t)
            c.appendChild(v)
        return c

    def clear(self, ref):
        c = self.find(ref)
        if c is None:
            return
        for tag in ("f", "v", "is"):
            for n in children(c, tag):
                c.removeChild(n)
        for a in ("t", "cm"):
            if c.hasAttribute(a):
                c.removeAttribute(a)

    def formula_node(self, ref):
        c = self.find(ref)
        return child(c, "f") if c is not None else None


def shared_strings(pkg):
    if "xl/sharedStrings.xml" not in pkg.raw:
        return []
    doc = pkg.xml("xl/sharedStrings.xml")
    return ["".join(text_of(t) for t in si.getElementsByTagName("t"))
            for si in doc.getElementsByTagName("si")]


def drop_part(pkg, part, rel_type_suffix):
    pkg.remove(part)
    rels = pkg.xml("xl/_rels/workbook.xml.rels")
    for r in rels.getElementsByTagName("Relationship"):
        if r.getAttribute("Type").endswith(rel_type_suffix):
            r.parentNode.removeChild(r)
    ct = pkg.xml("[Content_Types].xml")
    for o in ct.getElementsByTagName("Override"):
        if o.getAttribute("PartName") == "/" + part:
            o.parentNode.removeChild(o)


def prune_calc_chain(pkg, array_flags=False):
    """Keep the stale chain but drop entries whose cell no longer holds a formula (and,
    with `array_flags`, entries whose a="1" no longer matches the cell's formula kind).
    `i` is the sheetId and, when absent, repeats the previous entry's sheet."""
    wb = pkg.xml("xl/workbook.xml")
    by_id = {s.getAttribute("sheetId"): s.getAttribute("name") for s in wb.getElementsByTagName("sheet")}
    parts = pkg.sheet_parts()
    sheets = {n: Sheet(pkg.xml(parts[n])) for n in parts}
    chain = pkg.xml("xl/calcChain.xml")
    log, cur = [], None
    for c in list(chain.getElementsByTagName("c")):
        cur = c.getAttribute("i") or cur
        name = by_id[cur]
        cell = sheets[name].find(c.getAttribute("r"))
        f = child(cell, "f") if cell is not None else None
        stale_a = array_flags and f is not None and (c.getAttribute("a") == "1") != (f.getAttribute("t") == "array")
        if f is None or stale_a:
            nxt = c.nextSibling
            if nxt is not None and not nxt.getAttribute("i"):
                nxt.setAttribute("i", cur)   # the next entry inherited this sheet id
            c.parentNode.removeChild(c)
            log.append(f"calcChain: removed {name}!{c.getAttribute('r')} ({'array flag changed' if f is not None else 'no formula now'})")
    return log


def set_full_calc(pkg, on: bool):
    calc = pkg.xml("xl/workbook.xml").getElementsByTagName("calcPr")[0]
    if on:
        calc.setAttribute("fullCalcOnLoad", "1")
    elif calc.hasAttribute("fullCalcOnLoad"):
        calc.removeAttribute("fullCalcOnLoad")


def set_name(pkg, name, text):
    for d in pkg.xml("xl/workbook.xml").getElementsByTagName("definedName"):
        if d.getAttribute("name") == name and not d.hasAttribute("localSheetId"):
            while d.firstChild:
                d.removeChild(d.firstChild)
            d.appendChild(d.ownerDocument.createTextNode(text))
            return
    raise KeyError(name)


def dynamic_cm_index(pkg):
    """The 1-based cellMetadata index whose record points at XLDAPR fDynamic="1"."""
    if "xl/metadata.xml" not in pkg.raw:
        return None
    doc = pkg.xml("xl/metadata.xml")
    types = [t.getAttribute("name") for t in doc.getElementsByTagName("metadataType")]
    cm = doc.getElementsByTagName("cellMetadata")[0]
    for i, bk in enumerate(children(cm, "bk"), 1):
        rc = child(bk, "rc")
        if types[int(rc.getAttribute("t")) - 1] == "XLDAPR":
            return i
    return None


METADATA = ('<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
            '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" '
            'copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" '
            'clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
            '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
            '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
            '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>')


def add_metadata_part(pkg):
    pkg.add("xl/metadata.xml", XML_DECL + METADATA)
    rels = pkg.xml("xl/_rels/workbook.xml.rels")
    ids = {r.getAttribute("Id") for r in rels.getElementsByTagName("Relationship")}
    n = 1
    while f"rId{n}" in ids:
        n += 1
    r = rels.createElement("Relationship")
    r.setAttribute("Id", f"rId{n}")
    r.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata")
    r.setAttribute("Target", "metadata.xml")
    rels.documentElement.appendChild(r)
    ct = pkg.xml("[Content_Types].xml")
    o = ct.createElement("Override")
    o.setAttribute("PartName", "/xl/metadata.xml")
    o.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml")
    ct.documentElement.appendChild(o)


# ---- the patches ------------------------------------------------------------------
def patch_main(src, dst, drop_chain, keep_v, full_calc, calc_id=None):
    pkg = Package(src)
    log = []
    assert dynamic_cm_index(pkg) == 1, "expected cm=1 to be the dynamic-array record"
    sst = shared_strings(pkg)
    parts = pkg.sheet_parts()
    N, D, S, H = (Sheet(pkg.xml(parts[s])) for s in ("N", "D", "Slot", "Sh"))
    kv = dict(keep_v=keep_v, sst=sst)

    # 1 normal cells (and the cm traps)
    N.set_formula("B1", "SUM(A1:A3)*100", **kv)
    N.set_formula("B2", "_xlfn.LET(_xlpm.x,A1*10,_xlpm.x+1)", **kv)
    N.clear("B3")
    N.set_formula("B4", "SUM(A1:A3*2)", **kv)                 # trap: no cm
    N.set_formula("B5", "SUM(A1:A3*2)", array=True, **kv)     # as Excel stores it
    N.set_formula("B6", "1+2", array=True, **kv)              # scalar in array form
    N.set_formula("B7", "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(A3)", **kv)
    # 4 value and text cells
    N.set_formula("C1", "A1+100", **kv)
    N.set_formula("C2", 'UPPER("ab")&A1', **kv)
    N.set_formula("C3", "_xlfn.XLOOKUP(2,A1:A3,A1:A3)*7", **kv)
    N.set_formula("C4", "A1:A3*3", array=True, **kv)          # text cell -> spill

    # 2 spill anchors
    def anchor(ref, text, ghosts="delete", keep_ref=False, array=True):
        f = D.formula_node(ref)
        old = f.getAttribute("ref")
        D.set_formula(ref, text, array=array, **kv)
        if keep_ref:
            D.formula_node(ref).setAttribute("ref", old)
        if ghosts == "delete":
            for g in cells_in(old)[1:]:
                D.delete(g)
        log.append(f"D!{ref}: old ref {old} -> ref {D.formula_node(ref).getAttribute('ref') or '-'}, ghosts {ghosts}")
    anchor("C6", "_xlfn.SEQUENCE(5)")                          # bigger
    anchor("F6", "_xlfn.SEQUENCE(2)")                          # smaller
    anchor("H6", "_xlfn.SEQUENCE(5)", ghosts="keep", keep_ref=True)
    anchor("J6", "_xlfn.SEQUENCE(2)", ghosts="keep", keep_ref=True)
    anchor("L6", "_xlfn.SEQUENCE(2)", ghosts="keep")          # trap: ghosts left as constants
    anchor("N6", "42", array=False)                            # spill -> scalar
    D.set_formula("P6", "_xlfn.SEQUENCE(3)", array=True, **kv)  # scalar -> spill

    # 3 slots
    S.set_formula("B1", "_xlfn.SEQUENCE(1,3)*10", array=True, **kv)
    set_name(pkg, "Revenue", "_xlfn.ANCHORARRAY(Slot!$B$1)")
    S.set_formula("B2", "SUM(_xlfn.ANCHORARRAY(B1))/2", **kv)
    S.set_formula("B6", "Costs*0.25", **kv)
    S.set_formula("B9", "Tax+1", **kv)                         # row 9 is absent
    S.set_formula("F1", "SUM(Revenue)", **kv)

    # 5 shared formulas
    def group(col):
        f = H.formula_node(col + "1")
        return f.getAttribute("si"), f.getAttribute("ref")
    si, ref = group("B")
    H.set_formula("B3", "A3*100", **kv)                        # child: master untouched
    log.append(f"Sh!B3: child edited, master ref kept {ref}")
    si, ref = group("C")
    H.set_formula("C1", "A1+1000", **kv)                       # master: promote C2
    H.set_formula("C2", "A2+1", shared=(si, "C2:C5"), keep_v=True, sst=sst)
    log.append(f"Sh!C1: master edited, C2 promoted to master si={si} ref C2:C5 (was {ref})")
    si, ref = group("D")
    H.set_formula("D1", "A1+2000", **kv)                       # master: un-share the group
    for r in range(2, 6):
        H.set_formula(f"D{r}", f"A{r}*3", keep_v=True, sst=sst)
    log.append(f"Sh!D1: master edited, D2:D5 un-shared (group si={si} {ref} removed)")
    si, ref = group("E")
    H.set_formula("E5", "A5*400", **kv)                        # last child: shrink ref
    H.formula_node("E1").setAttribute("ref", "E1:E4")
    log.append(f"Sh!E5: last child edited, master ref {ref} -> E1:E4")

    # 6, 7 workbook-wide
    if drop_chain == "drop":
        drop_part(pkg, "xl/calcChain.xml", "/calcChain")
    elif drop_chain.startswith("prune"):
        log += prune_calc_chain(pkg, array_flags=drop_chain == "prune+a")
    set_full_calc(pkg, full_calc)
    if calc_id:
        pkg.xml("xl/workbook.xml").getElementsByTagName("calcPr")[0].setAttribute("calcId", calc_id)
    pkg.save(dst)
    return log


def patch_plain(src, dst, variant):
    pkg = Package(src)
    P = Sheet(pkg.xml(pkg.sheet_parts()["P"]))
    if variant == "q1":
        add_metadata_part(pkg)
        P.set_formula("B1", "_xlfn.SEQUENCE(3)*A1", array=True)
    else:
        P.set_formula("B1", "_xlfn.SEQUENCE(3)*A1", array="cse")
    P.set_formula("C1", "SUM(_xlfn.ANCHORARRAY(B1))")
    set_full_calc(pkg, True)
    pkg.save(dst)
    return [f"{variant}: metadata.xml {'created' if variant == 'q1' else 'absent'}; "
            f"B1 {'cm=1 t=array' if variant == 'q1' else 't=array, no cm'}"]


# ---- reading back ----------------------------------------------------------------
def stored(path: pathlib.Path):
    """Cell -> (t, cm, f attrs, f text, v) and the package facts."""
    pkg = Package(path)
    out = {"_parts": sorted(n for n in pkg.order if n.startswith("xl/") and n.count("/") == 1)}
    wb = pkg.xml("xl/workbook.xml")
    out["_calcPr"] = wb.getElementsByTagName("calcPr")[0].toxml()
    out["_names"] = {d.getAttribute("name"): text_of(d) for d in wb.getElementsByTagName("definedName")}
    if "xl/metadata.xml" in pkg.raw:
        out["_cellMetadata"] = pkg.xml("xl/metadata.xml").getElementsByTagName("cellMetadata")[0].toxml()
    if "xl/calcChain.xml" in pkg.raw:
        out["_calcChain"] = " ".join(c.getAttribute("r") + ("" if not c.getAttribute("i") else ":" + c.getAttribute("i"))
                                     for c in pkg.xml("xl/calcChain.xml").getElementsByTagName("c"))
    for sname, part in pkg.sheet_parts().items():
        sh = pkg.xml(part)
        out[f"_dimension {sname}"] = sh.getElementsByTagName("dimension")[0].getAttribute("ref")
        for c in sh.getElementsByTagName("c"):
            f, v = child(c, "f"), child(c, "v")
            fa = ""
            if f is not None:
                fa = " ".join(f"{k}={f.getAttribute(k)}" for k in ("t", "ref", "si") if f.hasAttribute(k))
            out[f"{sname}!{c.getAttribute('r')}"] = (
                c.getAttribute("t"), c.getAttribute("cm"), fa,
                text_of(f) if f is not None else None, text_of(v) if v is not None else None)
    return out


REPAIR_DIR = pathlib.Path.home() / "Library/Containers/com.microsoft.Excel/Data/tmp"


def repair_logs(stem, since):
    """Excel for Mac answers its own repair prompt when driven by AppleScript (the open
    goes on, repaired) but writes 'Repair Result to <stem><n>.xml' in its sandbox tmp."""
    out = []
    for p in REPAIR_DIR.glob(f"Repair Result to {stem}*.xml"):
        rest = p.stem[len("Repair Result to " + stem):]
        if rest.isdigit() and p.stat().st_mtime >= since:
            doc = minidom.parse(str(p))
            out += [text_of(r) for r in doc.getElementsByTagName("removedRecord")]
            out += [text_of(r) for r in doc.getElementsByTagName("repairedRecord")]
    return out


def parse_read(text):
    head, vals = "", {}
    for line in text.splitlines():
        if line.startswith("@"):
            ref, f, s = (line[1:].split("\t") + ["", ""])[:3]
            vals[ref] = (f, s)
        else:
            head += line + " "
    return head.strip(), vals


def fmt_cell(t):
    if t is None:
        return "(absent)"
    typ, cm, fa, f, v = t
    s = ""
    if typ:
        s += f"t={typ} "
    if cm:
        s += f"cm={cm} "
    if f is not None:
        s += f"<f {fa}>{f}" if fa else f"<f>{f}"
    elif fa:
        s += f"<f {fa}/>"
    if v is not None:
        s += f" v={v}"
    return s.strip() or "(empty element)"


def main():
    STAGE.mkdir(parents=True, exist_ok=True)
    for p in STAGE.glob("f8_*.xlsx"):
        p.unlink()
    was_running = subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True).returncode == 0
    if was_running:
        others = osa('tell application "Microsoft Excel" to get name of workbooks')
        print("Excel already running; open workbooks:", others or "(none)")
    rep = ["# excel-dim F8 — replace cell formulas at file level (macOS)",
           f"# {time.strftime('%FT%T')}", ""]
    rep.append("## 1 build: " + osa(BUILD, BASE, PLAIN))
    plain_oracle = STAGE / "f8_plain_oracle.xlsx"
    rep.append("## 2 oracle: " + osa(ORACLE_AS, BASE, ORACLE) + "; "
               + osa(PLAIN_ORACLE_AS, PLAIN, plain_oracle))

    rep.append("## 3 patches")
    files = {}
    for v, (drop_chain, keep_v, full_calc, calc_id) in VARIANTS.items():
        dst = STAGE / f"f8_{v}.xlsx"
        log = patch_main(BASE, dst, drop_chain, keep_v, full_calc, calc_id)
        files[v] = dst
        rep.append(f"   {v}: drop calcChain={drop_chain}, keep stale <v>={keep_v}, fullCalcOnLoad={full_calc}, calcId={calc_id or 'as saved'}")
        if v == "p1":
            rep += ["      " + l for l in log]
        elif v.startswith("p2"):
            rep += ["      " + l for l in log if l.startswith("calcChain")]
    for v in PLAIN_VARIANTS:
        dst = STAGE / f"f8_{v}.xlsx"
        rep += ["   " + l for l in patch_plain(PLAIN, dst, v)]
        files[v] = dst

    rep.append("## 4 Excel opens each file (values read before any recalculation request)")
    reads, resaved = {}, {}
    head, reads["oracle"] = parse_read(osa(READ_AS, ORACLE, "-", ",".join(CELLS), ORACLE.name))
    rep.append(f"   oracle: {head}")
    head, reads["plain_oracle"] = parse_read(osa(READ_AS, plain_oracle, "-", ",".join(PLAIN_CELLS), plain_oracle.name))
    rep.append(f"   plain oracle: {head}")
    # A stale calcChain brings up Excel's repair prompt, which blocks Excel until someone
    # answers it (System Events cannot press the button without Accessibility rights).
    # So the p2 variants are opened only with --with-repair, last, one attempt each.
    order = [v for v in files if not v.startswith("p2")]
    if "--with-repair" in sys.argv:
        order += [v for v in files if v.startswith("p2")]
    repaired = {}
    for v in order:
        path = files[v]
        out = STAGE / f"f8_{v}_resaved.xlsx"
        cells = PLAIN_CELLS if v.startswith("q") else CELLS
        t0 = time.time() - 1
        res = osa(READ_AS, path, out, ",".join(cells), path.name, timeout=90)
        head, reads[v] = parse_read(res)
        logs = repair_logs(path.stem, t0)
        if logs:
            repaired[v] = logs
            head += "  ** REPAIRED by Excel: " + " | ".join(logs)
        rep.append(f"   {v}: {head}")
        if not res.startswith("workbook "):
            rep.append(f"   {v}: REPAIR or no open: {res[:200]}")
            rep.append("   STOP: Excel is probably showing its repair prompt; nothing else is opened.")
            break
        if out.exists():
            resaved[v] = out

    rep.append("   => repaired by Excel: " + (", ".join(repaired) or "none"))
    if not was_running:
        osa('tell application "Microsoft Excel"\nif (count of workbooks) is 0 then quit\nend tell')

    # values on open vs oracle
    for group, oracle_key, names in (("main", "oracle", list(VARIANTS)), ("plain", "plain_oracle", list(PLAIN_VARIANTS))):
        cells = PLAIN_CELLS if group == "plain" else CELLS
        names = [n for n in names if n in reads]
        rep += ["", f"## 5a {group}: displayed values on open (oracle | " + " | ".join(names) + ")"]
        bad = {n: 0 for n in names}
        for ref in cells:
            o = reads[oracle_key].get(ref, ("?", "?"))
            row = []
            for n in names:
                got = reads[n].get(ref, ("?", "?"))
                ok = got[1] == o[1]
                bad[n] += not ok
                row.append(("" if ok else "!") + got[1])
            rep.append(f"   {ref:9} {o[1]:>8} | " + " | ".join(f"{x:>8}" for x in row) + f"    oracle f: {o[0]}")
        rep.append("   => values differing from the oracle: " + ", ".join(f"{n} {bad[n]}" for n in names))
        rep.append("   formula2 differences from the oracle on open:")
        for n in names:
            for ref in cells:
                o, got = reads[oracle_key].get(ref), reads[n].get(ref)
                if o and got and o[0] != got[0]:
                    rep.append(f"      {n} {ref}: oracle {o[0]!r} vs {got[0]!r}")

    # stored XML: oracle vs patched vs re-saved
    so, sp = stored(ORACLE), stored(files["p1"])
    rep += ["", "## 5b stored cells: oracle / p1 as patched / p1 re-saved by Excel (cells that differ anywhere)"]
    if "p1" in resaved:
        sr = stored(resaved["p1"])
        for k in sorted(set(so) | set(sp) | set(sr), key=lambda k: (k.startswith("_"), k)):
            vals = [so.get(k), sp.get(k), sr.get(k)]
            if k.startswith("_"):
                rep.append(f"   {k}:\n      oracle  {vals[0]}\n      patched {vals[1]}\n      resaved {vals[2]}")
                continue
            if k not in CELLS and vals[0] == vals[2]:
                continue
            o, p, r = (fmt_cell(x) for x in vals)
            mark = "same" if o == r else "DIFF"
            rep.append(f"   {mark} {k:9} oracle: {o:44} patched: {p:44} resaved: {r}")
    for v in [x for x in VARIANTS if x != "p1"] + list(PLAIN_VARIANTS):
        if v in resaved:
            s = stored(resaved[v])
            rep.append(f"   {v} re-saved: parts {s['_parts']}; calcPr {s['_calcPr']}; calcChain {s.get('_calcChain', '(none)')}")
    if "q1" in resaved:
        rep.append("   plain oracle stored: " + "; ".join(f"{k}={fmt_cell(v) if not k.startswith('_') else v}"
                                                         for k, v in stored(plain_oracle).items()))
        for v in PLAIN_VARIANTS:
            s = stored(resaved[v])
            rep.append(f"   {v} re-saved stored: " + "; ".join(f"{k}={fmt_cell(x) if not k.startswith('_') else x}"
                                                              for k, x in s.items()))

    for f in [BASE, PLAIN, ORACLE, plain_oracle, files["p1"], files["q1"], files["p2"], *resaved.values()]:
        if f.exists() and (f.name.startswith(("f8_base", "f8_plain_base", "f8_oracle", "f8_plain_oracle", "f8_p1", "f8_q1", "f8_p2."))):
            shutil.copy2(f, OUT / f.name)
    text = "\n".join(rep)
    (OUT / f"f8-mac-{time.strftime('%Y%m%d-%H%M%S')}.txt").write_text(text, encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
