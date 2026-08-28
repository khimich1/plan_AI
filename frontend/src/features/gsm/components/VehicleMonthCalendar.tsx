import type { CSSProperties } from "react";
import {
  layoutWeekSlots,
  type VehicleDayCell,
} from "@/features/gsm/lib/vehicleDayMap";
import type { GsmWaybill } from "@/features/gsm/types/gsm";

type Props = {
  cells: VehicleDayCell[];
  month?: string; // YYYY-MM
  onMonthChange?: (delta: number) => void;
  onDayClick?: (waybill: GsmWaybill) => void;
  onGapClick?: () => void;
};

const WEEKDAYS = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"] as const;

const GAP_LABEL = "нет путевого на заправку/мойку";

const wrap: CSSProperties = {
  display: "grid",
  gap: "0.5rem",
};

const grid: CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(7, minmax(0, 1fr))",
  gap: 4,
};

const headCell: CSSProperties = {
  textAlign: "center",
  fontSize: "0.7rem",
  color: "#667085",
  fontWeight: 600,
  padding: "0.15rem 0",
};

const slotStyle: CSSProperties = {
  minHeight: 44,
  borderRadius: 8,
  border: "1px solid transparent",
  background: "transparent",
};

const dayBase: CSSProperties = {
  minHeight: 44,
  borderRadius: 8,
  border: "1px solid #eaecf0",
  background: "#f9fafb",
  display: "flex",
  flexDirection: "column",
  alignItems: "center",
  justifyContent: "center",
  gap: 2,
  padding: "0.25rem 0.15rem",
  fontSize: "0.8rem",
  color: "#344054",
  width: "100%",
  boxSizing: "border-box",
};

const dayStyle = (cell: VehicleDayCell): CSSProperties => {
  if (cell.isGap) {
    return {
      ...dayBase,
      border: "1px solid #fdb022",
      background: "#fffaeb",
      cursor: "pointer",
    };
  }
  if (cell.isRed) {
    return {
      ...dayBase,
      border: "2px solid #f04438",
      background: "#fef3f2",
      cursor: cell.hasPl ? "pointer" : "default",
      boxShadow: "0 0 0 1px #f04438",
    };
  }
  if (cell.hasPl) {
    return { ...dayBase, cursor: "pointer" };
  }
  return dayBase;
};

const markerRow: CSSProperties = {
  display: "flex",
  gap: 3,
  fontSize: "0.65rem",
  lineHeight: 1,
  color: "#475467",
};

const dayNumber = (iso: string): string => String(Number(iso.slice(8, 10)));

const monthLabel = (month: string): string => {
  const raw = new Date(`${month}-01T12:00:00`).toLocaleDateString("ru-RU", {
    month: "long",
    year: "numeric",
  });
  return raw.charAt(0).toUpperCase() + raw.slice(1);
};

const headerStyle: CSSProperties = {
  display: "flex",
  alignItems: "center",
  justifyContent: "space-between",
  gap: "0.5rem",
};

const navButton: CSSProperties = {
  border: "1px solid #d0d5dd",
  borderRadius: 8,
  background: "#ffffff",
  color: "#344054",
  width: 32,
  height: 32,
  fontSize: "1rem",
  lineHeight: 1,
  cursor: "pointer",
};

const monthLabelStyle: CSSProperties = {
  fontSize: "0.95rem",
  fontWeight: 600,
  color: "#344054",
};

const ariaLabelFor = (cell: VehicleDayCell): string => {
  const day = dayNumber(cell.date);
  if (cell.isGap) return `${day}: ${GAP_LABEL}`;
  if (cell.isRed) return `${day}: ручная доработка`;
  if (cell.hasPl && cell.hasTx) return `${day}: ПЛ и транзакция`;
  if (cell.hasPl) return `${day}: путевой лист`;
  if (cell.hasTx) return `${day}: транзакция`;
  return day;
};

export const VehicleMonthCalendar = ({
  cells,
  month,
  onMonthChange,
  onDayClick,
  onGapClick,
}: Props) => {
  const slots = layoutWeekSlots(cells);
  const isEmptyPeriod =
    cells.length > 0 && cells.every((c) => !c.hasTx && !c.hasPl);

  return (
    <div style={wrap} data-testid="vehicle-month-calendar">
      {month && onMonthChange && (
        <div style={headerStyle}>
          <button
            type="button"
            aria-label="Предыдущий месяц"
            style={navButton}
            onClick={() => onMonthChange(-1)}
          >
            ‹
          </button>
          <span style={monthLabelStyle} data-testid="cal-month-label">
            {monthLabel(month)}
          </span>
          <button
            type="button"
            aria-label="Следующий месяц"
            style={navButton}
            onClick={() => onMonthChange(1)}
          >
            ›
          </button>
        </div>
      )}
      {isEmptyPeriod && (
        <p role="status" style={{ margin: 0, fontSize: "0.85rem", color: "#667085" }}>
          нет движений
        </p>
      )}
      <div style={grid} role="grid" aria-label="Календарь периода">
        {WEEKDAYS.map((label) => (
          <div key={label} style={headCell} role="columnheader">
            {label}
          </div>
        ))}
        {slots.map((slot, idx) => {
          if (slot === null) {
            return <div key={`pad-${idx}`} style={slotStyle} aria-hidden />;
          }

          const interactive = slot.isGap || slot.hasPl;
          const style = dayStyle(slot);
          const label = ariaLabelFor(slot);
          const markers = (
            <span style={markerRow} aria-hidden>
              {slot.hasTx && <span>tx</span>}
              {slot.hasPl && <span>ПЛ</span>}
            </span>
          );
          const body = (
            <>
              <span>{dayNumber(slot.date)}</span>
              {markers}
            </>
          );

          if (interactive) {
            return (
              <button
                key={slot.date}
                type="button"
                data-testid={`cal-day-${slot.date}`}
                aria-label={label}
                style={{ ...style, font: "inherit" }}
                onClick={() => {
                  if (slot.isGap) {
                    onGapClick?.();
                    return;
                  }
                  if (slot.waybill) onDayClick?.(slot.waybill);
                }}
              >
                {body}
              </button>
            );
          }

          return (
            <div
              key={slot.date}
              data-testid={`cal-day-${slot.date}`}
              aria-label={label}
              style={style}
            >
              {body}
            </div>
          );
        })}
      </div>
    </div>
  );
};
