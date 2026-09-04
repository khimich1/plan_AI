import { addDaysIso, type PromiseQuoteWindow } from "@/features/factory-capacity/api/promiseQuote";
import { isoWeekStart, toIsoDate } from "@/features/factory-capacity/lib/isoWeek";
import type { CSSProperties } from "react";

type Props = {
  month: string;
  minMonth: string;
  maxMonth: string;
  onMonthChange: (month: string) => void;
  selectedWeekStart: string | null;
  onSelectWeek: (weekStart: string) => void;
  onSelectDay?: (iso: string) => void;
  quoteWindow?: PromiseQuoteWindow | null;
  promisedDate?: string | null;
  firstPourDate?: string | null;
  pourFrom?: string | null;
  pourToSunday?: string | null;
  occupancy?: Record<string, number>;
  knob?: number;
  holidays?: string[];
  extraWorkdays?: string[];
};

const WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"];
const MONTHS = [
  "январь",
  "февраль",
  "март",
  "апрель",
  "май",
  "июнь",
  "июль",
  "август",
  "сентябрь",
  "октябрь",
  "ноябрь",
  "декабрь",
];

function parseMonth(iso: string): { year: number; monthIndex: number } {
  const [yearRaw, monthRaw] = iso.split("-");
  return { year: Number(yearRaw), monthIndex: Number(monthRaw) - 1 };
}

function isWeekend(iso: string): boolean {
  const day = new Date(`${iso}T12:00:00`).getDay();
  return day === 0 || day === 6;
}

function isNonWorking(
  iso: string,
  holidays: Set<string>,
  extraWorkdays: Set<string>,
): boolean {
  if (extraWorkdays.has(iso)) return false;
  if (holidays.has(iso)) return true;
  return isWeekend(iso);
}

function isPourDay(iso: string, pourFrom?: string | null, pourToSunday?: string | null): boolean {
  if (!pourFrom || !pourToSunday) return false;
  return iso >= pourFrom && iso <= pourToSunday;
}

function buildWeeks(month: string): string[][] {
  const { year, monthIndex } = parseMonth(month);
  const first = `${year}-${String(monthIndex + 1).padStart(2, "0")}-01`;
  const nextMonth = new Date(year, monthIndex + 1, 1);
  const last = toIsoDate(new Date(nextMonth.getTime() - 24 * 60 * 60 * 1000));
  const gridStart = isoWeekStart(first);
  const lastWeekStart = isoWeekStart(last);
  const weeks: string[][] = [];
  let cursor = gridStart;
  while (cursor <= lastWeekStart) {
    weeks.push(Array.from({ length: 7 }, (_, i) => addDaysIso(cursor, i)));
    cursor = addDaysIso(cursor, 7);
  }
  return weeks;
}

const wrapStyle: CSSProperties = {
  display: "grid",
  gap: "0.6rem",
};

const navStyle: CSSProperties = {
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  gap: "0.5rem",
};

const gridStyle: CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(7, 1fr)",
  gap: 2,
};

const weekdayStyle: CSSProperties = {
  textAlign: "center",
  fontSize: "0.7rem",
  color: "#667085",
  padding: "0.15rem 0",
};

const weekRowStyle: CSSProperties = {
  display: "contents",
};

const cellStyle = (opts: {
  outside: boolean;
  off: boolean;
  yellow: boolean;
  selected: boolean;
  overflow: boolean;
}): CSSProperties => ({
  position: "relative",
  border: "none",
  borderRadius: 6,
  minHeight: 38,
  padding: "2px 0 4px",
  fontSize: "0.8rem",
  cursor: "pointer",
  background: opts.overflow ? "#fef3f2" : opts.yellow ? "#fef9c3" : opts.selected ? "#eef4ff" : "transparent",
  color: opts.outside ? "#98a2b3" : opts.off ? "#98a2b3" : opts.overflow ? "#b42318" : "#101828",
  outline: opts.selected ? "2px solid #2e90fa" : opts.overflow ? "1px solid #fda29b" : "none",
  outlineOffset: -1,
});

export const PromisePeriodCalendar = ({
  month,
  minMonth,
  maxMonth,
  onMonthChange,
  selectedWeekStart,
  onSelectWeek,
  onSelectDay,
  promisedDate,
  firstPourDate,
  pourFrom,
  pourToSunday,
  occupancy = {},
  knob,
  holidays = [],
  extraWorkdays = [],
}: Props) => {
  const holidaySet = new Set(holidays);
  const extraSet = new Set(extraWorkdays);
  const { year, monthIndex } = parseMonth(month);
  const weeks = buildWeeks(month);
  const prevDisabled = month.slice(0, 7) <= minMonth.slice(0, 7);
  const nextDisabled = month.slice(0, 7) >= maxMonth.slice(0, 7);

  return (
    <section data-testid="promise-period-calendar" style={wrapStyle}>
      <div style={navStyle}>
        <button
          type="button"
          aria-label="Предыдущий месяц"
          disabled={prevDisabled}
          onClick={() => {
            const d = new Date(`${month}T12:00:00`);
            d.setMonth(d.getMonth() - 1);
            onMonthChange(monthStartSafe(d));
          }}
        >
          ‹
        </button>
        <strong style={{ textTransform: "capitalize" }}>
          {MONTHS[monthIndex]} {year}
        </strong>
        <button
          type="button"
          aria-label="Следующий месяц"
          disabled={nextDisabled}
          onClick={() => {
            const d = new Date(`${month}T12:00:00`);
            d.setMonth(d.getMonth() + 1);
            onMonthChange(monthStartSafe(d));
          }}
        >
          ›
        </button>
      </div>
      <div style={gridStyle} role="grid" aria-label="Календарь периодов">
        {WEEKDAYS.map((label) => (
          <div key={label} role="columnheader" style={weekdayStyle}>
            {label}
          </div>
        ))}
        {weeks.map((days) => {
          const weekStart = days[0];
          const selected = selectedWeekStart === weekStart;
          return (
            <div key={weekStart} role="row" data-week-start={weekStart} style={weekRowStyle}>
              {days.map((iso) => {
                const outside = iso.slice(0, 7) !== month.slice(0, 7);
                const off = isNonWorking(iso, holidaySet, extraSet);
                const yellow = isPourDay(iso, pourFrom, pourToSunday);
                const occupied = occupancy[iso] ?? 0;
                const overflow = Boolean(knob) && occupied > knob;
                const isPromised = promisedDate === iso;
                const isStart = firstPourDate === iso;
                const dayNum = Number(iso.slice(8, 10));
                const showFraction = Boolean(knob) && !off && occupied > 0;
                return (
                  <button
                    key={iso}
                    type="button"
                    role="gridcell"
                    data-testid={`promise-cal-day-${iso}`}
                    data-pour={yellow ? "true" : undefined}
                    data-overflow={overflow ? "true" : undefined}
                    aria-current={selected ? "true" : undefined}
                    onClick={() => {
                      onSelectWeek(isoWeekStart(iso));
                      onSelectDay?.(iso);
                    }}
                    style={cellStyle({ outside, off, yellow, selected, overflow })}
                  >
                    <span>{dayNum}</span>
                    {showFraction ? (
                      <span
                        data-testid={`promise-cal-frac-${iso}`}
                        style={{ display: "block", fontSize: "0.65rem", lineHeight: 1.1 }}
                      >
                        {occupied}/{knob}
                      </span>
                    ) : null}
                    {isStart ? (
                      <span
                        data-testid="promise-cal-start-marker"
                        title={`начало отливки ${iso}`}
                        style={{
                          display: "block",
                          width: 6,
                          height: 6,
                          borderRadius: "50%",
                          background: "#1570ef",
                          margin: "2px auto 0",
                        }}
                      />
                    ) : null}
                    {isPromised ? (
                      <span
                        data-testid="promise-cal-promised-marker"
                        title={`дата клиенту ${iso}`}
                        style={{
                          display: "block",
                          width: 6,
                          height: 6,
                          borderRadius: "50%",
                          background: "#92400e",
                          margin: "2px auto 0",
                        }}
                      />
                    ) : null}
                  </button>
                );
              })}
            </div>
          );
        })}
      </div>
    </section>
  );
};

function monthStartSafe(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-01`;
}
