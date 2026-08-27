import { useMemo, useState } from "react";
import type { CapacitySnapshot } from "@/features/factory-capacity/types/capacity";

const WEEK_DAYS = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];

type CellState = "empty" | "partial" | "full" | "holiday" | "outside" | "weekend";

function parseMonth(ym: string): Date {
  const [y, m] = ym.split("-").map(Number);
  return new Date(y, m - 1, 1);
}

function formatMonth(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function formatISO(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function addMonths(d: Date, n: number): Date {
  return new Date(d.getFullYear(), d.getMonth() + n, 1);
}

function monthLabel(d: Date): string {
  return d.toLocaleDateString("ru-RU", { month: "long", year: "numeric" });
}

function cellState(
  iso: string,
  snapshot: CapacitySnapshot,
  inMonth: boolean,
): CellState {
  if (!inMonth) return "outside";
  const holidays = new Set(snapshot.holidays);
  const extra = new Set(snapshot.extra_workdays);
  const d = new Date(iso + "T12:00:00");
  const weekend = d.getDay() === 0 || d.getDay() === 6;
  if (holidays.has(iso) || (weekend && !extra.has(iso))) {
    return weekend && !extra.has(iso) ? "weekend" : "holiday";
  }
  const info = snapshot.days_info[iso];
  const occupied = info?.occupied ?? 0;
  const max = info?.max ?? 5;
  if (occupied <= 0) return "empty";
  if (occupied >= max) return "full";
  return "partial";
}

const stateClass: Record<CellState, string> = {
  empty: "prod-calendar__day prod-calendar__day--empty",
  partial: "prod-calendar__day prod-calendar__day--partial",
  full: "prod-calendar__day prod-calendar__day--full",
  holiday: "prod-calendar__day prod-calendar__day--holiday",
  weekend: "prod-calendar__day prod-calendar__day--holiday",
  outside: "prod-calendar__day prod-calendar__day--outside",
};

type Props = {
  snapshot: CapacitySnapshot;
};

export const FactoryMiniCalendar = ({ snapshot }: Props) => {
  const from = useMemo(() => parseMonth(snapshot.calendar_from_month), [snapshot.calendar_from_month]);
  const to = useMemo(() => parseMonth(snapshot.calendar_to_month), [snapshot.calendar_to_month]);
  const [month, setMonth] = useState(() => from);

  // Clamp when snapshot range changes.
  const current = useMemo(() => {
    const key = formatMonth(month);
    if (key < snapshot.calendar_from_month) return from;
    if (key > snapshot.calendar_to_month) return to;
    return month;
  }, [month, from, to, snapshot.calendar_from_month, snapshot.calendar_to_month]);

  const canPrev = formatMonth(current) > snapshot.calendar_from_month;
  const canNext = formatMonth(current) < snapshot.calendar_to_month;

  const cells = useMemo(() => {
    const start = new Date(current.getFullYear(), current.getMonth(), 1);
    const end = new Date(current.getFullYear(), current.getMonth() + 1, 0);
    // Monday-based week
    const startPad = (start.getDay() + 6) % 7;
    const days: { iso: string; day: number; inMonth: boolean }[] = [];
    for (let i = 0; i < startPad; i++) {
      const d = new Date(start);
      d.setDate(d.getDate() - (startPad - i));
      days.push({ iso: formatISO(d), day: d.getDate(), inMonth: false });
    }
    for (let day = 1; day <= end.getDate(); day++) {
      const d = new Date(current.getFullYear(), current.getMonth(), day);
      days.push({ iso: formatISO(d), day, inMonth: true });
    }
    while (days.length % 7 !== 0) {
      const last = days[days.length - 1];
      const d = new Date(last.iso + "T12:00:00");
      d.setDate(d.getDate() + 1);
      days.push({ iso: formatISO(d), day: d.getDate(), inMonth: false });
    }
    return days;
  }, [current]);

  return (
    <div data-testid="factory-mini-calendar" style={{ display: "grid", gap: "0.5rem" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "0.5rem" }}>
        <button
          type="button"
          disabled={!canPrev}
          onClick={() => setMonth(addMonths(current, -1))}
          aria-label="Предыдущий месяц"
          style={{ opacity: canPrev ? 1 : 0.4 }}
        >
          ‹
        </button>
        <strong style={{ textTransform: "capitalize", fontSize: "0.9rem" }}>
          {monthLabel(current)}
        </strong>
        <button
          type="button"
          disabled={!canNext}
          onClick={() => setMonth(addMonths(current, 1))}
          aria-label="Следующий месяц"
          style={{ opacity: canNext ? 1 : 0.4 }}
        >
          ›
        </button>
      </div>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(7, 1fr)",
          gap: 4,
          fontSize: "0.7rem",
          textAlign: "center",
          color: "#667085",
        }}
      >
        {WEEK_DAYS.map((d) => (
          <div key={d}>{d}</div>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 4 }}>
        {cells.map((cell) => {
          const state = cellState(cell.iso, snapshot, cell.inMonth);
          return (
            <div
              key={cell.iso + (cell.inMonth ? "" : "-out")}
              className={stateClass[state]}
              title={cell.iso}
              style={{
                minHeight: 28,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: "0.75rem",
                borderRadius: 6,
                border: "1px solid #e4e7ec",
                pointerEvents: "none",
              }}
            >
              {cell.day}
            </div>
          );
        })}
      </div>
    </div>
  );
};
