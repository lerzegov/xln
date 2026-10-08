// Pretty-printer for long definitions (requirement A5) and comparison modulo whitespace.
// Whitespace in a formula is layout, except the space that is the intersection operator;
// Excel re-spaces definitions on entry (probe T14), so comparisons must ignore layout.
// It also rewrites number literals when it saves (`1E-14` comes back as
// `0.00000000000001`), so comparisons take a number for its value, not its spelling.

import { leftSpine, type Expr } from "./ast.js";
import { parse } from "./parser.js";
import { tokenize, type Token } from "./tokens.js";

export interface PrettyOptions {
  /** Target line width. Default 80. */
  width?: number;
  /** One level of indentation. Default four spaces. */
  indent?: string;
  /** Line break to use. Default "\n". */
  newline?: string;
  /**
   * Fill lines with the arguments of a long call when every argument is short (`MAX(C6,
   * C23, C28, …)`), instead of one argument per line. Default false.
   */
  pack?: boolean;
}

/** Arguments at most this long may share a line when `pack` is on. */
const PACK_ARG = 24;

const TIGHT_OPS = new Set([":", ":.", ".:", ".:.", "^"]);

function opText(op: string): string {
  if (op === " ") return " ";
  if (op === ",") return ", ";
  return TIGHT_OPS.has(op) ? op : " " + op + " ";
}

/**
 * Re-lays out a formula: flat when it fits in `width`, otherwise LET bindings one per line,
 * a LAMBDA body on its own lines, and long argument lists one argument per line.
 * The result is equal to the input modulo whitespace.
 */
export function prettyPrint(src: string, opts: PrettyOptions = {}): string {
  const width = opts.width ?? 80;
  const unit = opts.indent ?? "    ";
  const nl = opts.newline ?? "\n";
  const f = parse(src);
  const leaf = (n: { span: { start: number; end: number } }): string => src.slice(n.span.start, n.span.end);

  const flat = (n: Expr): string => {
    switch (n.kind) {
      case "missing":
        return "";
      case "paren":
        return "(" + flat(n.expr) + ")";
      case "unary":
        return n.op + flat(n.operand);
      case "postfix":
        return flat(n.operand) + n.op;
      case "binary": {
        const { first, links } = leftSpine(n);
        let out = flat(first);
        for (const b of links) out += opText(b.op) + flat(b.right);
        return out;
      }
      case "array":
        return "{" + n.rows.map((r) => r.map(flat).join(",")).join(";") + "}";
      case "call":
        return leaf(n.fn) + "(" + n.args.map(flat).join(", ") + ")";
      case "invoke":
        return flat(n.callee) + "(" + n.args.map(flat).join(", ") + ")";
      case "lambda":
        return leaf(n.fn) + "(" + [...n.params.map((p) => leaf(p)), flat(n.body)].join(", ") + ")";
      case "let":
        return leaf(n.fn) + "(" + [...n.bindings.flatMap((b) => [leaf(b.name), flat(b.value)]), flat(n.body)].join(", ") + ")";
      default:
        return leaf(n);
    }
  };

  const flatCache = new Map<Expr, string>();
  const flatOnce = (n: Expr): string => {
    let t = flatCache.get(n);
    if (t === undefined) flatCache.set(n, (t = flat(n)));
    return t;
  };

  /** `col` is the column where the node starts; `depth` the current indentation level. */
  const pp = (n: Expr, depth: number, col: number): string => {
    const one = flatOnce(n);
    if (col + one.length <= width && !one.includes("\n")) return one;
    const ind = unit.repeat(depth);
    const ind1 = unit.repeat(depth + 1);
    const block = (head: string, items: string[], tail = ")"): string =>
      head + nl + items.map((s) => ind1 + s).join("," + nl) + nl + ind + tail;
    switch (n.kind) {
      case "let": {
        const items = n.bindings.map((b) => {
          const name = leaf(b.name);
          return name + ", " + pp(b.value, depth + 1, ind1.length + name.length + 2);
        });
        items.push(pp(n.body, depth + 1, ind1.length));
        return block(leaf(n.fn) + "(", items);
      }
      case "lambda": {
        const params = n.params.map((p) => leaf(p));
        const head = leaf(n.fn) + "(" + params.map((p) => p + ",").join(" ");
        return head + nl + ind1 + pp(n.body, depth + 1, ind1.length) + nl + ind + ")";
      }
      case "call": {
        if (n.args.length === 0) return one;
        const flats = n.args.map(flatOnce);
        if (opts.pack && flats.every((a) => a.length <= PACK_ARG && !a.includes("\n"))) {
          const rows: string[] = [];
          let row = "";
          for (const a of flats) {
            if (row !== "" && ind1.length + row.length + 2 + a.length + 1 > width) {
              rows.push(row + ",");
              row = a;
            } else row = row === "" ? a : row + ", " + a;
          }
          rows.push(row);
          return leaf(n.fn) + "(" + nl + rows.map((r) => ind1 + r).join(nl) + nl + ind + ")";
        }
        return block(leaf(n.fn) + "(", n.args.map((a) => pp(a, depth + 1, ind1.length)));
      }
      case "invoke": {
        const callee = pp(n.callee, depth, col);
        const lastLine = callee.slice(callee.lastIndexOf("\n") + 1);
        const args = n.args.map(flat).join(", ");
        if (lastLine.length + args.length + 2 <= width) return callee + "(" + args + ")";
        return block(callee + "(", n.args.map((a) => pp(a, depth + 1, ind1.length)));
      }
      case "paren": {
        const inner = pp(n.expr, depth, col + 1);
        return "(" + inner + ")";
      }
      case "binary": {
        // Operands stay on one line; only calls inside them break.
        const { first, links } = leftSpine(n);
        let out = pp(first, depth, col);
        for (const b of links) {
          const op = opText(b.op);
          const lastLine = out.length - out.lastIndexOf("\n") - 1;
          out += op + pp(b.right, depth, lastLine + op.length);
        }
        return out;
      }
      case "unary":
        return n.op + pp(n.operand, depth, col + n.op.length);
      case "postfix":
        return pp(n.operand, depth, col) + n.op;
      default:
        return one;
    }
  };

  const lead = f.equals ? "=" : "";
  return lead + pp(f.body, 0, lead.length);
}

// ---------------------------------------------------------------------------------------
// Comparison
// ---------------------------------------------------------------------------------------

/** Removes layout spaces between the items of a structured reference: `[[#This Row], [A]]`. */
function canonicalStructInner(inner: string): string {
  // `T[Col A]`: a single bare column name, where every space is part of the name.
  if (!inner.includes("[")) return inner;
  let out = "";
  let depth = 0;
  for (let k = 0; k < inner.length; k++) {
    const c = inner[k]!;
    if (c === "'" && depth > 0) {
      out += c + (inner[k + 1] ?? "");
      k++;
      continue;
    }
    if (c === "[") depth++;
    else if (c === "]") depth--;
    if (depth === 0 && (c === " " || c === "\t" || c === "\r" || c === "\n")) continue;
    out += c;
  }
  return out;
}

/**
 * The canonical spelling of a number literal (the tokenizer's `number`: digits, decimals,
 * exponent, never a sign): its value as JavaScript prints it, after rounding to the 15
 * significant digits Excel keeps of a typed number. `1E-14`, `1e-14` and
 * `0.00000000000001` give `1e-14`; `.5` and `0.50` give `0.5`; `1.0` and `1` give `1`.
 * Measured: Excel saves `1E-14` as `0.00000000000001`. Assumed: that it rounds, not
 * truncates, past 15 digits (probes/README.md, "Number literals").
 */
export function canonicalNumber(text: string): string {
  const n = Number(text);
  if (!Number.isFinite(n)) return text;
  return String(Number(n.toPrecision(15)));
}

export interface TokenKeyOptions {
  /** `text`: numbers as written (the lockfile hash of formats 1 and 2). Default `value`. */
  numbers?: "value" | "text";
}

function tokenKey(t: Token, opts: TokenKeyOptions): string {
  switch (t.kind) {
    case "number":
      return opts.numbers === "text" ? `number:${t.text}` : `number:${canonicalNumber(t.text)}`;
    case "isect":
      return "isect";
    case "structref":
      return `structref:${t.value ?? ""}[${canonicalStructInner(t.inner ?? "")}]`;
    case "ref":
    case "name": {
      if (!t.qual) return `${t.kind}:${t.value}`;
      // `'BS'!A1` and `BS!A1` are the same reference: compare the sheet, not its quoting.
      const q = t.qual;
      return `${t.kind}:${q.book ?? ""}|${q.sheet ?? ""}|${q.sheet2 ?? ""}!${t.value}`;
    }
    default:
      return `${t.kind}:${t.text}`;
  }
}

/** The tokens of a formula that carry meaning, as comparable keys. */
export function tokenKeys(src: string, opts: TokenKeyOptions = {}): string[] {
  return tokenize(src)
    .filter((t) => t.kind !== "ws" && t.kind !== "eof")
    .map((t) => tokenKey(t, opts));
}

/**
 * Whether two formulas are the same modulo layout whitespace (and optional quotes around
 * sheet names, and the spelling of number literals: `1E-14` is `0.00000000000001`). The
 * space that is the intersection operator still counts.
 */
export function equalModuloWhitespace(a: string, b: string): boolean {
  const ka = tokenKeys(a);
  const kb = tokenKeys(b);
  return ka.length === kb.length && ka.every((k, i) => k === kb[i]);
}

/** Converts line breaks in layout whitespace to CR LF (how Excel stores them); strings are untouched. */
export function toCrLf(src: string): string {
  return tokenize(src)
    .map((t) => (t.kind === "ws" || t.kind === "isect" ? t.text.replace(/\r?\n/g, "\r\n") : t.text))
    .join("");
}
