/** 1-based index for non-empty (trim) lines; blank → null. */
export function assignNonEmptyLineNumbers(lines: readonly string[]): Array<number | null> {
  let n = 0;
  return lines.map((line) => {
    if (!line.trim()) {
      return null;
    }
    n += 1;
    return n;
  });
}

const LINE_NUMBER_SUFFIX = ". ";

/** Visual gutter label: `1. `, `2. `. Not written into the list text. */
export function formatPlateLineNumber(n: number): string {
  return `${n}${LINE_NUMBER_SUFFIX}`;
}

/** Gutter width in `ch` for the formatted label (`1. `, `10. `, …). */
export function lineNumberGutterCh(maxNumber: number): number {
  return formatPlateLineNumber(Math.max(maxNumber, 1)).length;
}
