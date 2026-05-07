/** Парсинг срока изготовления КП — зеркало `core/execution_terms.py` (даты, затем «N дней» / «N недель»). */

/** Подсказка под полем (единая для мастера КП и архива). */
export const EXECUTION_TERMS_FIELD_HINT =
  "Формат: ДД.ММ.ГГГГ, ГГГГ-ММ-ДД, «7 дней» или «2 недели». Можно ввести календарную дату или срок от сегодня.";

export const EXECUTION_TERMS_PLACEHOLDER = "Например, 14 дней или 05.06.2026";

export const EXECUTION_TERMS_PARSE_ERROR =
  "Не удалось распознать срок. Используйте формат ДД.ММ.ГГГГ, ГГГГ-ММ-ДД, «N дней» или «N недель».";

const DATE_DD_MM_YYYY = /^(\d{2})\.(\d{2})\.(\d{4})$/;
const DATE_ISO = /^(\d{4})-(\d{2})-(\d{2})$/;
const DAYS_RE = /(\d+)\s*(?:дн|день|дней|day|days)/i;
const WEEKS_RE = /(\d+)\s*(?:нед|недел|недели|week|weeks)/i;

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function parseAbsoluteDate(trimmed: string): string | null {
  let m = DATE_DD_MM_YYYY.exec(trimmed);
  if (m) {
    const dd = Number(m[1]);
    const mm = Number(m[2]);
    const yyyy = Number(m[3]);
    const d = new Date(yyyy, mm - 1, dd);
    if (d.getFullYear() === yyyy && d.getMonth() === mm - 1 && d.getDate() === dd) {
      return `${pad2(dd)}.${pad2(mm)}.${yyyy}`;
    }
    return null;
  }
  m = DATE_ISO.exec(trimmed);
  if (m) {
    const yyyy = Number(m[1]);
    const mm = Number(m[2]);
    const dd = Number(m[3]);
    const d = new Date(yyyy, mm - 1, dd);
    if (d.getFullYear() === yyyy && d.getMonth() === mm - 1 && d.getDate() === dd) {
      return `${pad2(dd)}.${pad2(mm)}.${yyyy}`;
    }
    return null;
  }
  return null;
}

function addDays(base: Date, days: number): Date {
  const out = new Date(base);
  out.setDate(out.getDate() + days);
  return out;
}

function addWeeks(base: Date, weeks: number): Date {
  return addDays(base, weeks * 7);
}

function formatDdMmYyyy(d: Date): string {
  return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()}`;
}

/** Возвращает нормализованную дату `ДД.ММ.ГГГГ` или `null`. */
export function tryNormalizeExecutionTerms(raw: string, now: Date = new Date()): string | null {
  const trimmed = (raw ?? "").trim();
  if (!trimmed) return null;

  const absolute = parseAbsoluteDate(trimmed);
  if (absolute) return absolute;

  const daysM = DAYS_RE.exec(trimmed);
  if (daysM) {
    const n = parseInt(daysM[1], 10);
    if (Number.isFinite(n) && n >= 0) {
      return formatDdMmYyyy(addDays(now, n));
    }
  }

  const weeksM = WEEKS_RE.exec(trimmed);
  if (weeksM) {
    const n = parseInt(weeksM[1], 10);
    if (Number.isFinite(n) && n >= 0) {
      return formatDdMmYyyy(addWeeks(now, n));
    }
  }

  return null;
}
