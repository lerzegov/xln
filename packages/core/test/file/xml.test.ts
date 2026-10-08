import { describe, expect, it } from "vitest";
import { attrNS, decodeEntities, parseXml, XmlError, XmlReader, type XmlToken } from "../../src/file/xml.js";
import { resolveTarget } from "../../src/file/package.js";
import { columnName, parseCell } from "../../src/file/cellref.js";
import { isLocked, lockFileName } from "../../src/file/lock.js";

function tokens(src: string): XmlToken[] {
  const r = new XmlReader(src);
  const out: XmlToken[] = [];
  for (let t = r.next(); t; t = r.next()) out.push(t);
  return out;
}

describe("xml reader", () => {
  it("keeps CR LF in text content as stored", () => {
    const root = parseXml('<?xml version="1.0"?>\r\n<a>x,\r\n  y</a>');
    expect(root.children).toEqual(["x,\r\n  y"]);
  });

  it("decodes entities and character references", () => {
    expect(decodeEntities("a&lt;b&amp;&quot;&apos;&gt;&#65;&#x3bb;")).toBe("a<b&\"'>Aλ");
    expect(() => decodeEntities("&nbsp;")).toThrow(XmlError);
  });

  it("normalises literal whitespace in attributes but keeps references", () => {
    const root = parseXml('<a v="x\r\ny\tz&#10;w"/>');
    expect(root.attrs["v"]).toBe("x y z\nw");
  });

  it("merges CDATA with text and skips comments and PIs", () => {
    const root = parseXml("<a>1<![CDATA[<2>]]><!-- c --><?pi x?>3</a>");
    expect(root.children).toEqual(["1<2>3"]);
  });

  it("reports self-closing elements as open + close", () => {
    expect(tokens("<a><b x='1'/></a>").map((t) => t.type)).toEqual(["open", "open", "close", "close"]);
  });

  it("resolves namespaces of elements and attributes", () => {
    const root = parseXml('<w xmlns="urn:m" xmlns:r="urn:r"><s r:id="rId1"/><x:t xmlns:x="urn:x"/></w>');
    expect(root.ns).toBe("urn:m");
    const [s, t] = root.children as [ReturnType<typeof parseXml>, ReturnType<typeof parseXml>];
    expect(attrNS(s, ["urn:r"], "id")).toBe("rId1");
    expect(attrNS(s, ["urn:other"], "id")).toBeUndefined();
    expect(t.ns).toBe("urn:x");
    expect(t.local).toBe("t");
  });

  it("rejects malformed documents", () => {
    for (const bad of ["<a><b></a>", "<a>", "<a x=1/>", "</a>", "<a/><b/>", "<p:a/>", "<a x='1' x='2'/>", "<a>&bogus;</a>"]) {
      expect(() => parseXml(bad), bad).toThrow(XmlError);
    }
  });

  it("skips a subtree", () => {
    const r = new XmlReader("<r><a><b>t</b><c/></a><d/></r>");
    r.next(); // r
    const a = r.next();
    if (a?.type !== "open") throw new Error();
    r.skip(a);
    const d = r.next();
    expect(d?.type === "open" && d.local).toBe("d");
  });
});

describe("package paths", () => {
  it("resolves relative, parent and absolute targets", () => {
    expect(resolveTarget("xl", "worksheets/sheet1.xml")).toBe("xl/worksheets/sheet1.xml");
    expect(resolveTarget("xl/worksheets", "../tables/table1.xml")).toBe("xl/tables/table1.xml");
    expect(resolveTarget("xl", "/xl/worksheets/sheet2.xml")).toBe("xl/worksheets/sheet2.xml");
    expect(resolveTarget("", "xl/workbook.xml")).toBe("xl/workbook.xml");
    expect(resolveTarget("xl", "./a%20b.xml")).toBe("xl/a b.xml");
  });
});

describe("cell addresses", () => {
  it("round-trips columns and parses addresses", () => {
    expect(columnName(1)).toBe("A");
    expect(columnName(26)).toBe("Z");
    expect(columnName(27)).toBe("AA");
    expect(columnName(16384)).toBe("XFD");
    expect(parseCell("XFD1048576")).toEqual({ row: 1048576, col: 16384 });
    expect(parseCell("$B$3")).toEqual({ row: 3, col: 2 });
    expect(parseCell("B")).toBeUndefined();
    expect(parseCell("3")).toBeUndefined();
    expect(parseCell("B0")).toBeUndefined();
  });
});

describe("lock file", () => {
  it("names and detects Excel's owner file", () => {
    expect(lockFileName("model.xlsx")).toBe("~$model.xlsx");
    expect(lockFileName("C:\\work\\model.xlsx")).toBe("~$model.xlsx");
    expect(lockFileName("/Users/x/model.xlsx")).toBe("~$model.xlsx");
    expect(isLocked("model.xlsx", ["model.xlsx", "~$model.xlsx"])).toBe(true);
    expect(isLocked("model.xlsx", ["model.xlsx", "~$MODEL.XLSX"])).toBe(true);
    expect(isLocked("model.xlsx", ["model.xlsx", "~$other.xlsx"])).toBe(false);
    expect(isLocked("/a/model.xlsx", ["/a/~$model.xlsx"])).toBe(true);
  });
});
