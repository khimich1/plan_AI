import {
  canAddDayToBasket,
  getDayKind,
  type BasketDayKind,
  type DayKind,
} from "@/features/production/lib/basketDayKind";
import type { DayInfo, FillTargetItem } from "@/features/production/types/production";

const MIX_ERROR =
  "Нельзя смешивать свободные и частично занятые дни. Часть дней не добавлена.";

const SKIP_ERROR =
  "Часть дней пропущена (выходные, заполненные или выполненные).";

const parseISODate = (iso: string): Date => {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d);
};

const formatISODate = (d: Date): string => {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
};

const isWeekend = (d: Date): boolean => d.getDay() === 0 || d.getDay() === 6;

const resolveDayInfo = (
  date: string,
  daysInfo: Record<string, DayInfo>,
  defaultMax: number,
): DayInfo =>
  daysInfo[date] ?? {
    occupied: 0,
    max: defaultMax,
    completed: false,
    day_number: 0,
  };

export type BrushSelectOptions = {
  iso: string;
  holidays: Set<string>;
  extraWorkdays: Set<string>;
  defaultMax?: number;
};

/** Inclusive ISO date range, sorted ascending (local calendar days). */
export function datesBetweenInclusive(a: string, b: string): string[] {
  const start = parseISODate(a <= b ? a : b);
  const end = parseISODate(a <= b ? b : a);
  const out: string[] = [];
  for (let cur = new Date(start); cur <= end; cur.setDate(cur.getDate() + 1)) {
    out.push(formatISODate(cur));
  }
  return out;
}

/** Workday with free slots, not completed — eligible for brush paint. */
export function isDayBrushSelectable(
  dayInfo: DayInfo | undefined,
  options: BrushSelectOptions,
): boolean {
  const { iso, holidays, extraWorkdays, defaultMax = 5 } = options;
  const info = dayInfo ?? {
    occupied: 0,
    max: defaultMax,
    completed: false,
    day_number: 0,
  };
  if (info.completed) return false;
  const date = parseISODate(iso);
  if (holidays.has(iso) || (isWeekend(date) && !extraWorkdays.has(iso))) {
    return false;
  }
  const freeSlots = Math.max(0, info.max - info.occupied);
  return freeSlots > 0;
}

export type PaintDaysArgs = {
  dates: string[];
  brushTracks: number;
  daysInfo: Record<string, DayInfo>;
  basketKind: BasketDayKind | null;
  holidays?: Set<string>;
  extraWorkdays?: Set<string>;
  defaultMax?: number;
};

export type PaintDaysResult = {
  added: FillTargetItem[];
  error: string | null;
};

/**
 * Paint dates with brushTracks (clamped to freeSlots).
 * Skips non-selectable / kind-incompatible days and sets error if any skipped.
 */
export function paintDays(args: PaintDaysArgs): PaintDaysResult {
  const {
    dates,
    brushTracks,
    daysInfo,
    basketKind,
    holidays = new Set(),
    extraWorkdays = new Set(),
    defaultMax = 5,
  } = args;

  const added: FillTargetItem[] = [];
  let currentKind: BasketDayKind | null = basketKind;
  let skippedNonSelectable = false;
  let skippedKindMix = false;

  for (const date of dates) {
    const info = resolveDayInfo(date, daysInfo, defaultMax);
    if (
      !isDayBrushSelectable(daysInfo[date], {
        iso: date,
        holidays,
        extraWorkdays,
        defaultMax,
      })
    ) {
      skippedNonSelectable = true;
      continue;
    }

    const dayKind: DayKind = getDayKind(info);
    if (!canAddDayToBasket(currentKind, dayKind)) {
      skippedKindMix = true;
      continue;
    }

    const freeSlots = Math.max(0, info.max - info.occupied);
    const tracks = Math.max(1, Math.min(brushTracks, freeSlots));
    added.push({ date, tracks });

    if (currentKind === null && (dayKind === "empty" || dayKind === "partial")) {
      currentKind = dayKind;
    }
  }

  let error: string | null = null;
  if (skippedKindMix) {
    error = MIX_ERROR;
  } else if (skippedNonSelectable) {
    error = SKIP_ERROR;
  }

  return { added, error };
}
