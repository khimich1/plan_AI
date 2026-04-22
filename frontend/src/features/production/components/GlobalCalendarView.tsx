import { useMemo, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { Spinner } from "@/shared/ui/Spinner";
import { DayDrawer } from "@/features/production/components/DayDrawer";
import {
  useGlobalCalendarQuery,
  useWorkCalendarQuery,
} from "@/features/production/hooks/useProductionQueries";
import type { DayInfo } from "@/features/production/types/production";

const WEEK_DAYS = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];

type CellState = "empty" | "partial" | "full" | "completed" | "holiday" | "outside";

const stateClass: Record<CellState, string> = {
  empty: "prod-calendar__day prod-calendar__day--empty",
  partial: "prod-calendar__day prod-calendar__day--partial",
  full: "prod-calendar__day prod-calendar__day--full",
  completed: "prod-calendar__day prod-calendar__day--completed",
  holiday: "prod-calendar__day prod-calendar__day--holiday",
  outside: "prod-calendar__day prod-calendar__day--outside",
};

const formatISO = (d: Date) => {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
};

const addDays = (d: Date, days: number) => {
  const next = new Date(d);
  next.setDate(next.getDate() + days);
  return next;
};

const startOfMonth = (d: Date) => new Date(d.getFullYear(), d.getMonth(), 1);
const endOfMonth = (d: Date) => new Date(d.getFullYear(), d.getMonth() + 1, 0);

const isWeekend = (d: Date) => d.getDay() === 0 || d.getDay() === 6;

export const GlobalCalendarView = () => {
  const [month, setMonth] = useState(() => startOfMonth(new Date()));
  const [selectedDate, setSelectedDate] = useState<string | null>(null);

  const calendarQuery = useGlobalCalendarQuery();
  const workCalendar = useWorkCalendarQuery();

  const daysInfo = calendarQuery.data?.days_info ?? {};
  const holidays = useMemo(
    () => new Set(workCalendar.data?.extra_holidays ?? []),
    [workCalendar.data],
  );
  const extraWorkdays = useMemo(
    () => new Set(workCalendar.data?.extra_workdays ?? []),
    [workCalendar.data],
  );

  const monthStart = startOfMonth(month);
  const monthEnd = endOfMonth(month);
  const gridStartOffset = (monthStart.getDay() + 6) % 7; // понедельник = 0
  const totalCells = Math.ceil((gridStartOffset + monthEnd.getDate()) / 7) * 7;
  const gridStart = addDays(monthStart, -gridStartOffset);

  const buildCellState = (date: Date, info: DayInfo | undefined): CellState => {
    const iso = formatISO(date);
    const isOutside = date.getMonth() !== month.getMonth();
    if (isOutside) return "outside";
    if (info?.completed) return "completed";
    if (info && info.occupied >= info.max) return "full";
    if (info && info.occupied > 0) return "partial";
    if (holidays.has(iso) || (isWeekend(date) && !extraWorkdays.has(iso))) {
      return "holiday";
    }
    return "empty";
  };

  const cells = Array.from({ length: totalCells }, (_, index) => {
    const date = addDays(gridStart, index);
    const iso = formatISO(date);
    const info = daysInfo[iso];
    const state = buildCellState(date, info);
    return { date, iso, info, state };
  });

  const isLoading = calendarQuery.isLoading || workCalendar.isLoading;

  const selectedInfo = selectedDate ? daysInfo[selectedDate] : undefined;

  return (
    <Card
      title="Календарный план"
      subtitle="Сводная загрузка всех планов по датам. Клик по дню — чтобы посмотреть содержимое, скачать документы и отметить выполнение."
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: "1rem",
        }}
      >
        <div style={{ display: "flex", gap: "0.5rem" }}>
          <Button
            variant="secondary"
            onClick={() => setMonth(addDays(month, -1 * month.getDate()))}
          >
            ← Месяц
          </Button>
          <Button variant="secondary" onClick={() => setMonth(startOfMonth(new Date()))}>
            Сегодня
          </Button>
          <Button
            variant="secondary"
            onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))}
          >
            Месяц →
          </Button>
        </div>
        <strong style={{ fontSize: "1.1rem" }}>
          {month.toLocaleDateString("ru-RU", { month: "long", year: "numeric" })}
        </strong>
      </div>

      {isLoading && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}>
          <Spinner /> Загрузка календаря…
        </div>
      )}

      {calendarQuery.isError && (
        <Alert tone="error">Не удалось загрузить календарь производства.</Alert>
      )}

      {!isLoading && (
        <div className="prod-calendar">
          <div className="prod-calendar__week-header">
            {WEEK_DAYS.map((w) => (
              <div key={w} className="prod-calendar__week-header-cell">
                {w}
              </div>
            ))}
          </div>
          <div className="prod-calendar__grid">
            {cells.map(({ date, iso, info, state }) => {
              const max = info?.max ?? 5;
              const occupied = info?.occupied ?? 0;
              const disabled = state === "outside";
              return (
                <button
                  key={iso}
                  type="button"
                  disabled={disabled}
                  className={stateClass[state]}
                  onClick={() => {
                    if (!disabled) setSelectedDate(iso);
                  }}
                >
                  <span className="prod-calendar__date-number">{date.getDate()}</span>
                  {info && (
                    <span className="prod-calendar__occupancy">
                      {occupied}/{max}
                    </span>
                  )}
                  {info?.completed && <span className="prod-calendar__badge">✓</span>}
                </button>
              );
            })}
          </div>
          <div className="prod-calendar__legend">
            <span className="prod-calendar__legend-item prod-calendar__day--empty">Свободен</span>
            <span className="prod-calendar__legend-item prod-calendar__day--partial">Частично</span>
            <span className="prod-calendar__legend-item prod-calendar__day--full">Заполнен</span>
            <span className="prod-calendar__legend-item prod-calendar__day--completed">Выполнен</span>
            <span className="prod-calendar__legend-item prod-calendar__day--holiday">Выходной</span>
          </div>
        </div>
      )}

      <DayDrawer
        date={selectedDate}
        summary={selectedInfo}
        onClose={() => setSelectedDate(null)}
      />
    </Card>
  );
};
