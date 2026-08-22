import { describe, expect, it } from "vitest";

import {
  assignNonEmptyLineNumbers,
  formatPlateLineNumber,
  lineNumberGutterCh,
} from "@/features/commercial-offer/lib/plateListLineNumbers";

describe("assignNonEmptyLineNumbers", () => {
  it("returns empty array for no lines", () => {
    expect(assignNonEmptyLineNumbers([])).toEqual([]);
  });

  it("returns null for a single empty string", () => {
    expect(assignNonEmptyLineNumbers([""])).toEqual([null]);
  });

  it("treats whitespace-only lines as empty", () => {
    expect(assignNonEmptyLineNumbers(["   ", "\t"])).toEqual([null, null]);
  });

  it("numbers two non-empty lines as 1 and 2", () => {
    expect(assignNonEmptyLineNumbers(["ПБ 78-12-8н 2", "71-12-8 3"])).toEqual([1, 2]);
  });

  it("skips a blank line between plates without consuming a number", () => {
    expect(assignNonEmptyLineNumbers(["A", "", "B"])).toEqual([1, null, 2]);
  });

  it("skips leading and trailing empty lines from split", () => {
    expect(assignNonEmptyLineNumbers(["", "A", "B", ""])).toEqual([null, 1, 2, null]);
  });

  it("counts a trimmed non-empty line", () => {
    expect(assignNonEmptyLineNumbers(["  A  ", "B"])).toEqual([1, 2]);
  });
});

describe("formatPlateLineNumber", () => {
  it("adds a period and space after the digit", () => {
    expect(formatPlateLineNumber(1)).toBe("1. ");
    expect(formatPlateLineNumber(12)).toBe("12. ");
  });
});

describe("lineNumberGutterCh", () => {
  it("fits `9. ` in 3ch", () => {
    expect(lineNumberGutterCh(9)).toBe(3);
  });

  it("grows with two-digit labels", () => {
    expect(lineNumberGutterCh(10)).toBe(4);
    expect(lineNumberGutterCh(99)).toBe(4);
  });

  it("grows to 5ch at 100", () => {
    expect(lineNumberGutterCh(100)).toBe(5);
  });
});
