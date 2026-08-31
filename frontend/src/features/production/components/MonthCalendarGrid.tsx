import { useEffect, useRef, useState } from "react";
import { Button } from "@/shared/ui/Button";
import type {
  CalendarViewMode,
  DayInfo,
} from "@/features/production/types/production";

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

const MODE_OPTIONS: { value: CalendarViewMode; label: string }[] = [
  { value: "planning", label: "Планирование" },
  { value: "capacity", label: "Ёмкость" },
];

/** Hard factory cap: day max tracks cannot exceed this. */
const TRACKS_PER_DAY_HARD_CAP = 5;

const clampCapacityTracks = (value: number): number =>
  Math.max(0, Math.min(TRACKS_PER_DAY_HARD_CAP, value));

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

export type DayActivateMeta = {
  shiftKey: boolean;
};

export type MonthCalendarGridProps = {
  daysInfo: Record<string, DayInfo>;
  holidays: Set<string>;
  extraWorkdays: Set<string>;
  month: Date;
  onMonthChange: (next: Date) => void;
  selectedDate: string | null;
  /**
   * Кисть: клик / Shift+клик. Не открывает drawer —
   * для просмотра дня используйте onOpenDay.
   */
  onDayActivate: (iso: string, meta: DayActivateMeta) => void;
  /** Двойной клик или кнопка «i» — открыть DayDrawer. */
  onOpenDay?: (iso: string) => void;
  /** @deprecated Используйте onDayActivate. */
  onSelectDate?: (iso: string) => void;
  /** Дни, добавленные в корзину — выделяем рамкой/чек-меткой. */
  highlightedDates?: Set<string>;
  /** Число дорожек в корзине по дате (бейдж на ячейке). */
  basketTracksByDate?: Record<string, number>;
  /** Режим сетки; если не передан — внутренний toggle. */
  mode?: CalendarViewMode;
  onModeChange?: (mode: CalendarViewMode) => void;
  /** max_tracks по дате (режим ёмкости). */
  dayCapacity?: Record<string, number>;
  /** Сохранить override ёмкости. */
  onSaveDayCapacity?: (date: string, maxTracks: number) => void | Promise<void>;
  capacitySaving?: boolean;
};

/**
 * Презентационный компонент: сетка месяца + переключатели + легенда.
 *
 * Жесты (планирование): клик → onDayActivate; Shift+клик → range meta;
 * double-click / «i» → onOpenDay.
 * Режим ёмкости: кисть отключена; клик по числу max → inline edit.
 */
export const MonthCalendarGrid = ({
  daysInfo,
  holidays,
  extraWorkdays,
  month,
  onMonthChange,
  selectedDate,
  onDayActivate,
  onOpenDay,
  onSelectDate,
  highlightedDates,
  basketTracksByDate,
  mode: modeProp,
  onModeChange,
  dayCapacity,
  onSaveDayCapacity,
  capacitySaving = false,
}: MonthCalendarGridProps) => {
  const [internalMode, setInternalMode] = useState<CalendarViewMode>("planning");
  const mode = modeProp ?? internalMode;
  const setMode = (next: CalendarViewMode) => {
    if (modeProp === undefined) setInternalMode(next);
    onModeChange?.(next);
  };

  const [editingDate, setEditingDate] = useState<string | null>(null);
  const [draftTracks, setDraftTracks] = useState(0);
  const monthStart = startOfMonth(month);
  const monthEnd = endOfMonth(month);
  const gridStartOffset = (monthStart.getDay() + 6) % 7;
  const totalCells = Math.ceil((gridStartOffset + monthEnd.getDate()) / 7) * 7;
  const gridStart = addDays(monthStart, -gridStartOffset);
  // Откладываем single-click, чтобы double-click не успел toggle+toggle.
  const clickTimerRef = useRef<number | null>(null);
  const isCapacity = mode === "capacity";

  useEffect(() => {
    return () => {
      if (clickTimerRef.current != null) {
        window.clearTimeout(clickTimerRef.current);
      }
    };
  }, []);

  useEffect(() => {
    setEditingDate(null);
  }, [mode, month]);

  const buildCellState = (date: Date, info: DayInfo | undefined): CellState => {
    const iso = formatISO(date);
    const isOutside = date.getMonth() !== month.getMonth();
    if (isOutside) return "outside";
    if (isCapacity) return "empty";
    if (info?.completed) return "completed";
    if (info && info.occupied >= info.max) return "full";
    if (info && info.occupied > 0) return "partial";
    if (holidays.has(iso) || (isWeekend(date) && !extraWorkdays.has(iso))) {
      return "holiday";
    }
    return "empty";
  };

  const activate = (iso: string, shiftKey: boolean) => {
    onDayActivate(iso, { shiftKey });
    // Back-compat for callers still using onSelectDate.
    onSelectDate?.(iso);
  };

  const beginEdit = (iso: string, current: number) => {
    setEditingDate(iso);
    setDraftTracks(clampCapacityTracks(current));
  };

  const commitEdit = async () => {
    if (!editingDate || !onSaveDayCapacity) {
      setEditingDate(null);
      return;
    }
    const value = clampCapacityTracks(Math.floor(Number(draftTracks)));
    if (!Number.isFinite(value)) {
      setEditingDate(null);
      return;
    }
    await onSaveDayCapacity(editingDate, value);
    setEditingDate(null);
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
        role="tablist"
        aria-label="Режим календаря"
        className="prod-calendar__mode-toggle"
      >
        {MODE_OPTIONS.map((option) => {
          const isActive = mode === option.value;
          return (
            <button
              key={option.value}
              role="tab"
              type="button"
              aria-selected={isActive}
              className={
                isActive
                  ? "prod-calendar__mode-btn prod-calendar__mode-btn--active"
                  : "prod-calendar__mode-btn"
              }
              onClick={() => setMode(option.value)}
            >
              {option.label}
            </button>
          );
        })}
      </div>

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
            const isHighlighted = !isCapacity && (highlightedDates?.has(iso) ?? false);
            const isSelected = !isCapacity && selectedDate === iso;
            const basketTracks = basketTracksByDate?.[iso];
            const capacityValue = dayCapacity?.[iso] ?? max;
            const className = [
              isCapacity && !disabled
                ? "prod-calendar__day prod-calendar__day--capacity"
                : stateClass[state],
              isHighlighted ? "prod-calendar__day--in-basket" : "",
              isSelected ? "prod-calendar__day--selected" : "",
            ]
              .filter(Boolean)
              .join(" ");
            const isEditing = isCapacity && editingDate === iso;
            const capacityBody =
              isCapacity && !disabled ? (
                isEditing ? (
                  <span className="prod-calendar__capacity-edit">
                    <button
                      type="button"
                      className="prod-calendar__capacity-step"
                      aria-label="Уменьшить ёмкость"
                      disabled={capacitySaving || draftTracks <= 0}
                      onClick={() => setDraftTracks((v) => Math.max(0, v - 1))}
                    >
                      −
                    </button>
                    <input
                      type="number"
                      min={0}
                      max={TRACKS_PER_DAY_HARD_CAP}
                      step={1}
                      className="prod-calendar__capacity-input"
                      aria-label={`Ёмкость ${iso}`}
                      value={draftTracks}
                      disabled={capacitySaving}
                      onChange={(e) => {
                        const next = Number(e.target.value);
                        setDraftTracks(
                          Number.isFinite(next) ? clampCapacityTracks(next) : 0,
                        );
                      }}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          e.preventDefault();
                          void commitEdit();
                        }
                        if (e.key === "Escape") {
                          e.preventDefault();
                          setEditingDate(null);
                        }
                      }}
                      autoFocus
                    />
                    <button
                      type="button"
                      className="prod-calendar__capacity-step"
                      aria-label="Увеличить ёмкость"
                      disabled={
                        capacitySaving || draftTracks >= TRACKS_PER_DAY_HARD_CAP
                      }
                      onClick={() =>
                        setDraftTracks((v) => clampCapacityTracks(v + 1))
                      }
                    >
                      +
                    </button>
                    <button
                      type="button"
                      className="prod-calendar__capacity-save"
                      aria-label="Сохранить ёмкость"
                      disabled={capacitySaving}
                      onClick={() => void commitEdit()}
                    >
                      ✓
                    </button>
                  </span>
                ) : (
                  <button
                    type="button"
                    className="prod-calendar__capacity-value"
                    aria-label={`Изменить ёмкость ${iso}`}
                    onClick={() => beginEdit(iso, capacityValue)}
                  >
                    {capacityValue}
                  </button>
                )
              ) : null;

            return (
              <div key={iso} className="prod-calendar__day-wrap">
                {isCapacity ? (
                  <div className={className} aria-disabled={disabled || undefined}>
                    <span className="prod-calendar__date-number">{date.getDate()}</span>
                    {capacityBody}
                  </div>
                ) : (
                  <button
                    type="button"
                    disabled={disabled}
                    className={className}
                    onClick={(e) => {
                      if (disabled) return;
                      const shiftKey = e.shiftKey;
                      if (clickTimerRef.current != null) {
                        window.clearTimeout(clickTimerRef.current);
                        clickTimerRef.current = null;
                      }
                      // Shift+range сразу; plain click ждёт dblclick, если есть onOpenDay.
                      if (onOpenDay && !shiftKey) {
                        clickTimerRef.current = window.setTimeout(() => {
                          clickTimerRef.current = null;
                          activate(iso, false);
                        }, 250);
                        return;
                      }
                      activate(iso, shiftKey);
                    }}
                    onDoubleClick={(e) => {
                      e.preventDefault();
                      if (clickTimerRef.current != null) {
                        window.clearTimeout(clickTimerRef.current);
                        clickTimerRef.current = null;
                      }
                      if (disabled || !onOpenDay) return;
                      onOpenDay(iso);
                    }}
                  >
                    <span className="prod-calendar__date-number">{date.getDate()}</span>
                    {info && (
                      <span className="prod-calendar__occupancy">
                        {occupied}/{max}
                      </span>
                    )}
                    {info?.completed && <span className="prod-calendar__badge">✓</span>}
                    {isHighlighted && basketTracks != null && (
                      <span
                        className="prod-calendar__basket-tracks"
                        aria-label={`${basketTracks} дорожек в корзине`}
                      >
                        {basketTracks} дор.
                      </span>
                    )}
                    {isHighlighted && basketTracks == null && (
                      <span className="prod-calendar__basket-mark" aria-label="В корзине">
                        ●
                      </span>
                    )}
                  </button>
                )}
                {!disabled && !isCapacity && onOpenDay ? (
                  <button
                    type="button"
                    className="prod-calendar__day-info"
                    aria-label={`Открыть день ${iso}`}
                    title="Подробнее о дне"
                    onClick={(e) => {
                      e.stopPropagation();
                      onOpenDay(iso);
                    }}
                  >
                    i
                  </button>
                ) : null}
              </div>
            );
          })}
        </div>
        <div className="prod-calendar__legend">
          {isCapacity ? (
            <span className="prod-calendar__legend-item prod-calendar__day--capacity">
              Ёмкость (дорожек/день)
            </span>
          ) : (
            <>
              <span className="prod-calendar__legend-item prod-calendar__day--empty">
                Свободен
              </span>
              <span className="prod-calendar__legend-item prod-calendar__day--partial">
                Частично
              </span>
              <span className="prod-calendar__legend-item prod-calendar__day--full">
                Заполнен
              </span>
              <span className="prod-calendar__legend-item prod-calendar__day--completed">
                Выполнен
              </span>
              <span className="prod-calendar__legend-item prod-calendar__day--holiday">
                Выходной
              </span>
            </>
          )}
        </div>
      </div>
    </>
  );
};
