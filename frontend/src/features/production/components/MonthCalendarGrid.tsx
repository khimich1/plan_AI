import { Button } from "@/shared/ui/Button";
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

export type MonthCalendarGridProps = {
  daysInfo: Record<string, DayInfo>;
  holidays: Set<string>;
  extraWorkdays: Set<string>;
  month: Date;
  onMonthChange: (next: Date) => void;
  selectedDate: string | null;
  onSelectDate: (iso: string) => void;
  /** Дни, добавленные в корзину дозаполнения — выделяем рамкой/чек-меткой. */
  highlightedDates?: Set<string>;
};

/**
 * Презентационный компонент: сетка месяца + переключатели + легенда.
 *
 * Не знает ни про корзину, ни про DayDrawer — это ответственность
 * родителя. Выделение `highlightedDates` отдельным классом позволяет
 * использовать одну и ту же сетку и в `GlobalCalendarView`, и на шаге 1
 * мастера создания плана.
 */
export const MonthCalendarGrid = ({
  daysInfo,
  holidays,
  extraWorkdays,
  month,
  onMonthChange,
  selectedDate,
  onSelectDate,
  highlightedDates,
}: MonthCalendarGridProps) => {
  const monthStart = startOfMonth(month);
  const monthEnd = endOfMonth(month);
  const gridStartOffset = (monthStart.getDay() + 6) % 7;
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

  return (
    <>
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
            onClick={() => onMonthChange(addDays(month, -1 * month.getDate()))}
          >
            ← Месяц
          </Button>
          <Button variant="secondary" onClick={() => onMonthChange(startOfMonth(new Date()))}>
            Сегодня
          </Button>
          <Button
            variant="secondary"
            onClick={() => onMonthChange(new Date(month.getFullYear(), month.getMonth() + 1, 1))}
          >
            Месяц →
          </Button>
        </div>
        <strong style={{ fontSize: "1.1rem" }}>
          {month.toLocaleDateString("ru-RU", { month: "long", year: "numeric" })}
        </strong>
      </div>

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
            const isHighlighted = highlightedDates?.has(iso) ?? false;
            const isSelected = selectedDate === iso;
            const className = [
              stateClass[state],
              isHighlighted ? "prod-calendar__day--in-basket" : "",
              isSelected ? "prod-calendar__day--selected" : "",
            ]
              .filter(Boolean)
              .join(" ");
            return (
              <button
                key={iso}
                type="button"
                disabled={disabled}
                className={className}
                onClick={() => {
                  if (!disabled) onSelectDate(iso);
                }}
              >
                <span className="prod-calendar__date-number">{date.getDate()}</span>
                {info && (
                  <span className="prod-calendar__occupancy">
                    {occupied}/{max}
                  </span>
                )}
                {info?.completed && <span className="prod-calendar__badge">✓</span>}
                {isHighlighted && (
                  <span className="prod-calendar__basket-mark" aria-label="В корзине">
                    ●
                  </span>
                )}
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
    </>
  );
};
