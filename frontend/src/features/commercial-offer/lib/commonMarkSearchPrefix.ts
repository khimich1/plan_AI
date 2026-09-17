const TRAILING_PUNCT_RE = /[.\-\s]+$/u;
const LENGTH_SUFFIX_RE = /-\d+(?:\.\d+)?$/u;

/** C↔С / T↔Т / B↔В, collapse spaces, case-fold — same idea as catalog search. */
export const normalizeMarkForSearch = (mark: string): string =>
  mark
    .trim()
    .toUpperCase()
    .replace(/Ё/g, "Е")
    .replace(/С/g, "C")
    .replace(/В/g, "B")
    .replace(/Т/g, "T")
    .replace(/\s+/g, "");

const longestCommonPrefix = (left: string, right: string): string => {
  const limit = Math.min(left.length, right.length);
  let index = 0;
  while (index < limit && left[index] === right[index]) {
    index += 1;
  }
  return left.slice(0, index);
};

/** Shared search prefill: LCP of normalized marks, then a family stem, length ≥ 4. */
export const commonMarkSearchPrefix = (marks: string[]): string => {
  const normalized = marks.map(normalizeMarkForSearch).filter((item) => item.length > 0);
  if (normalized.length === 0) {
    return "";
  }
  let prefix = normalized[0] ?? "";
  for (const mark of normalized.slice(1)) {
    prefix = longestCommonPrefix(prefix, mark);
    if (prefix.replace(TRAILING_PUNCT_RE, "").length < 4) {
      return "";
    }
  }
  prefix = prefix.replace(TRAILING_PUNCT_RE, "");
  // One missing mark (C110.30-6) must not prefill the exact hole: catalog
  // lookup would be empty and hide neighbors such as C110.30-9.
  if (LENGTH_SUFFIX_RE.test(prefix)) {
    const stemmed = prefix.replace(LENGTH_SUFFIX_RE, "");
    if (stemmed.length >= 4) {
      prefix = stemmed;
    }
  }
  return prefix.length >= 4 ? prefix : "";
};

export const catalogRowSharesUnpricedStem = (catalogMark: string, missingMarks: string[]): boolean => {
  const catalog = normalizeMarkForSearch(catalogMark);
  if (!catalog) {
    return false;
  }
  return missingMarks.some((missing) => {
    const other = normalizeMarkForSearch(missing);
    const shared = longestCommonPrefix(catalog, other).replace(TRAILING_PUNCT_RE, "");
    return shared.length >= 4;
  });
};
