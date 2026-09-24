export type LongPileTariffMap = Record<string, { trip_cost: number }>;

export type LongPileTariffRow = {
  length_key: number;
  trip_cost?: number | null;
};

export function uniqueMarks(marks: Array<string | null | undefined>): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const mark of marks) {
    if (typeof mark !== "string" || mark.length === 0 || seen.has(mark)) {
      continue;
    }
    seen.add(mark);
    result.push(mark);
  }
  return result;
}

/** 140 → «14,0 м», 138 → «13,8 м». */
export function formatLongPileMeters(lengthKey: number): string {
  return `${(lengthKey / 10).toFixed(1).replace(".", ",")} м`;
}

export function parseTariffDraft(raw: string): number | null {
  const normalized = raw.trim().replace(/\s+/g, "").replace(",", ".");
  if (!normalized.length) {
    return null;
  }
  const value = Number(normalized);
  return Number.isFinite(value) ? value : null;
}

/**
 * Пустое поле не пишется и не становится 0: если тариф уже был, он сохраняется.
 * Явный 0 попадает в JSON.
 */
export function collectLongPileTariffs(
  lengths: LongPileTariffRow[],
  drafts: Record<string, string>,
): { ok: true; tariffs: LongPileTariffMap } | { ok: false; error: string } {
  const tariffs: LongPileTariffMap = {};
  for (const row of lengths) {
    const key = String(row.length_key);
    const raw = (drafts[key] ?? "").trim();
    if (!raw.length) {
      if (row.trip_cost !== null && row.trip_cost !== undefined) {
        tariffs[key] = { trip_cost: row.trip_cost };
      }
      continue;
    }
    const parsed = parseTariffDraft(raw);
    if (parsed === null || parsed < 0) {
      return { ok: false, error: "Тариф рейса должен быть числом не меньше 0." };
    }
    tariffs[key] = { trip_cost: parsed };
  }
  return { ok: true, tariffs };
}
