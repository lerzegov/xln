// Offsets ↔ 0-based line/character, as VS Code counts them (UTF-16 units; CR LF, LF and
// CR each end a line). Files that are not open in an editor have no TextDocument to ask.

export interface LinePos {
  line: number;
  character: number;
}

export class LineIndex {
  private readonly starts: number[] = [0];

  constructor(readonly text: string) {
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      if (c === 10) this.starts.push(i + 1);
      else if (c === 13) {
        if (text.charCodeAt(i + 1) === 10) i++;
        this.starts.push(i + 1);
      }
    }
  }

  get lineCount(): number {
    return this.starts.length;
  }

  position(offset: number): LinePos {
    const o = Math.max(0, Math.min(offset, this.text.length));
    let lo = 0;
    let hi = this.starts.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.starts[mid]! <= o) lo = mid;
      else hi = mid - 1;
    }
    return { line: lo, character: o - this.starts[lo]! };
  }

  offset(pos: LinePos): number {
    const start = this.starts[Math.max(0, Math.min(pos.line, this.starts.length - 1))]!;
    return Math.min(start + pos.character, this.text.length);
  }

  /** The text of a 0-based line, without its line break. */
  lineText(line: number): string {
    const s = this.starts[line] ?? this.text.length;
    const e = this.starts[line + 1] ?? this.text.length;
    return this.text.slice(s, e).replace(/\r?\n$|\r$/, "");
  }
}
