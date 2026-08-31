/** Convert ДД.ММ.ГГГГ → YYYY-MM-DD for capacity-snapshot target. */
export function ddMmYyyyToIso(value: string): string | null {
  const m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(value.trim());
  if (!m) return null;
  const dd = Number(m[1]);
  const mm = Number(m[2]);
  const yyyy = Number(m[3]);
  const d = new Date(yyyy, mm - 1, dd);
  if (d.getFullYear() !== yyyy || d.getMonth() !== mm - 1 || d.getDate() !== dd) {
    return null;
  }
  const pad = (n: number) => (n < 10 ? `0${n}` : String(n));
  return `${yyyy}-${pad(mm)}-${pad(dd)}`;
}

/** Max ISO produce_by from a list (invalid/empty skipped). */
export function maxIsoDate(dates: string[]): string | null {
  let best: string | null = null;
  for (const raw of dates) {
    const v = raw.trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) continue;
    if (best === null || v > best) best = v;
  }
  return best;
}
