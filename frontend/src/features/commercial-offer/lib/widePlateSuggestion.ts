/** Suggests standard-width replacements for plates wider than 12 dm. */
export const buildAutoSplitSuggestion = (line: string, fallbackQty: number): string => {
  const normalized = line.trim().replace(",", ".");
  const lineMatch = normalized.match(/^(.*\S)\s+(\d+)$/);
  const platePart = (lineMatch?.[1] ?? normalized).trim();
  const qty = lineMatch?.[2] ?? String(fallbackQty > 0 ? fallbackQty : 1);
  const nameMatch = platePart.match(/^(?:ПБ\s+)?([\d.]+)-([\d.]+)-(.+)$/i);
  if (!nameMatch) {
    return `ПБ 60-12-8п ${qty}\nПБ 60-3.0-8п ${qty}`;
  }

  const lengthRaw = nameMatch[1];
  const widthRaw = nameMatch[2];
  const suffix = nameMatch[3];
  const prefix = `ПБ ${lengthRaw}`;
  const widthDm = Number(widthRaw);
  if (!Number.isFinite(widthDm) || widthDm <= 12) {
    return `${prefix}-12-${suffix} ${qty}\n${prefix}-3.0-${suffix} ${qty}`;
  }

  const remainder = Math.max(widthDm - 12, 0);
  const remainderFormatted = remainder.toFixed(1);
  return `${prefix}-12-${suffix} ${qty}\n${prefix}-${remainderFormatted}-${suffix} ${qty}`;
};
