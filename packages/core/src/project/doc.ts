// Doc comments of LAMBDAs (M3c): `/** Summary. @param x what x is @param y … */`. AFE
// writes plain `/** */` comments; the `@param` tags are xln's addition, read for hover and
// signature help. The whole text, tags included, is the Name Manager comment, so a pull
// gives it back as written.

export interface DocParam {
  /** As written after `@param`, brackets of an optional parameter removed. */
  name: string;
  text: string;
}

export interface DocComment {
  /** The text before the first `@param`, trimmed. */
  summary: string;
  params: DocParam[];
}

const TAG = "@param";

function isSpace(c: string | undefined): boolean {
  return c === " " || c === "\t" || c === "\n" || c === "\r";
}

/** Splits a doc comment's text into its summary and its `@param` entries. */
export function parseDocComment(doc: string): DocComment {
  // A tag starts at the beginning or after whitespace, and is followed by whitespace.
  const starts: number[] = [];
  for (let i = doc.indexOf(TAG); i >= 0; i = doc.indexOf(TAG, i + 1)) {
    if ((i === 0 || isSpace(doc[i - 1])) && isSpace(doc[i + TAG.length])) starts.push(i);
  }
  const summary = (starts.length ? doc.slice(0, starts[0]) : doc).trim();
  const params: DocParam[] = [];
  starts.forEach((s, k) => {
    const body = doc.slice(s + TAG.length, k + 1 < starts.length ? starts[k + 1] : doc.length);
    let i = 0;
    while (isSpace(body[i])) i++;
    let j = i;
    while (j < body.length && !isSpace(body[j])) j++;
    let name = body.slice(i, j);
    if (name.startsWith("[") && name.endsWith("]")) name = name.slice(1, -1);
    if (name === "") return;
    const text = body
      .slice(j)
      .split("\n")
      .map((l) => l.trim())
      .filter((l) => l !== "")
      .join(" ");
    params.push({ name, text: text.startsWith("- ") ? text.slice(2) : text });
  });
  return { summary, params };
}

/** The documentation of a parameter (`x` or `[x]`), matched case-insensitively. */
export function paramDoc(doc: DocComment, param: string): string | undefined {
  const p = param.startsWith("[") && param.endsWith("]") ? param.slice(1, -1) : param;
  return doc.params.find((d) => d.name.toLowerCase() === p.toLowerCase())?.text;
}

/** An `@param` of a doc comment as it stands in a file, for the checks and their fixes. */
export interface DocParamSpan {
  /** As written, brackets of an optional parameter removed. */
  name: string;
  /** Whether it was written `[name]`. */
  bracketed: boolean;
  /** Offsets in the text scanned: the `@param` tag, and the name (inside the brackets). */
  tag: number;
  nameStart: number;
  nameEnd: number;
  /**
   * What removing it takes out: from the end of what precedes it (the previous line, or the
   * text before it on its line) to the same point of the next `@param`, or to the end of its
   * own text when it is the last.
   */
  removeStart: number;
  removeEnd: number;
}

/**
 * The `@param` tags of a doc comment in a file's text: `text.slice(start, end)` is what lies
 * between `/**` and `*\/`. Read like `parseDocComment` (a tag after whitespace, followed by
 * whitespace), with offsets into `text`.
 */
export function docParamSpans(text: string, start: number, end: number): DocParamSpan[] {
  const tags: number[] = [];
  for (let i = text.indexOf(TAG, start); i >= 0 && i + TAG.length <= end; i = text.indexOf(TAG, i + 1)) {
    if ((i === start || isSpace(text[i - 1])) && (i + TAG.length === end || isSpace(text[i + TAG.length]))) tags.push(i);
  }
  // Where removing a tag starts: back over the blanks before it on its line, and over a
  // continuation `*` and the line break when nothing else precedes it there.
  const cut = (i: number): number => {
    let k = i;
    while (k > start && (text[k - 1] === " " || text[k - 1] === "\t")) k--;
    let j = k;
    if (j > start && text[j - 1] === "*") {
      let m = j - 1;
      while (m > start && (text[m - 1] === " " || text[m - 1] === "\t")) m--;
      if (m === start || text[m - 1] === "\n") j = m;
    }
    if (j > start && text[j - 1] === "\n") return text[j - 2] === "\r" ? j - 2 : j - 1;
    return k;
  };
  const out: DocParamSpan[] = [];
  tags.forEach((t, k) => {
    let i = t + TAG.length;
    while (i < end && (text[i] === " " || text[i] === "\t")) i++;
    let j = i;
    while (j < end && !isSpace(text[j])) j++;
    if (j === i) return;
    let nameStart = i;
    let nameEnd = j;
    const bracketed = text[i] === "[" && text[j - 1] === "]" && j - i > 2;
    if (bracketed) {
      nameStart++;
      nameEnd--;
    }
    let removeEnd: number;
    if (k + 1 < tags.length) removeEnd = cut(tags[k + 1]!);
    else {
      removeEnd = end;
      // The last one runs to the end of its text, not over the comment's closing layout.
      while (removeEnd > j && isSpace(text[removeEnd - 1])) removeEnd--;
    }
    out.push({ name: text.slice(nameStart, nameEnd), bracketed, tag: t, nameStart, nameEnd, removeStart: cut(t), removeEnd });
  });
  return out;
}
