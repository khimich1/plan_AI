import type { CSSProperties } from "react";
import {
  addDaysIso,
  formatQuoteDayMonth,
  type PromiseQuoteWeek,
  type PromiseQuoteWindow,
} from "@/features/factory-capacity/api/promiseQuote";

type Props = {
  weeks: PromiseQuoteWeek[];
  quoteWindow?: PromiseQuoteWindow | null;
};

const stripStyle: CSSProperties = {
  display: "flex",
  gap: "0.5rem",
  overflowX: "auto",
  paddingBottom: "0.25rem",
};

const cardStyle: CSSProperties = {
  flex: "0 0 auto",
  minWidth: 132,
  display: "grid",
  gap: "0.35rem",
  padding: "0.65rem 0.75rem",
  border: "1px solid #e4e7ec",
  borderRadius: 10,
  background: "#ffffff",
  fontSize: "0.8rem",
};

const cardInWindowStyle: CSSProperties = {
  ...cardStyle,
  borderColor: "#98a2b3",
  background: "#f9fafb",
};

const rowStyle: CSSProperties = {
  display: "flex",
  justifyContent: "space-between",
  gap: "0.5rem",
  color: "#475467",
};

function weekLabel(weekStart: string): string {
  const weekEnd = addDaysIso(weekStart, 6);
  return `${formatQuoteDayMonth(weekStart)}–${formatQuoteDayMonth(weekEnd)}`;
}

function isInWindow(weekStart: string, quoteWindow: PromiseQuoteWindow | null | undefined): boolean {
  if (!quoteWindow) return false;
  return weekStart >= quoteWindow.from_week && weekStart <= quoteWindow.to_week;
}

export const PromiseWeekStrip = ({ weeks, quoteWindow }: Props) => {
  if (weeks.length === 0) {
    return (
      <div data-testid="promise-week-strip" role="status" style={{ fontSize: "0.85rem", color: "#667085" }}>
        Нет данных по неделям
      </div>
    );
  }

  return (
    <div data-testid="promise-week-strip" role="list" style={stripStyle}>
      {weeks.map((week) => {
        const inWindow = isInWindow(week.week_start, quoteWindow);
        return (
          <article
            key={week.week_start}
            role="listitem"
            aria-current={inWindow ? "true" : undefined}
            style={inWindow ? cardInWindowStyle : cardStyle}
          >
            <strong style={{ color: "#101828" }}>{weekLabel(week.week_start)}</strong>
            <div style={rowStyle}>
              <span>план</span>
              <span>{week.planned}</span>
            </div>
            <div style={rowStyle}>
              <span>обещано</span>
              <span>{week.promised}</span>
            </div>
            <div style={rowStyle}>
              <span>холды</span>
              <span>{week.held}</span>
            </div>
            <div style={rowStyle}>
              <span>свободно</span>
              <strong style={{ color: "#101828" }}>{week.free}</strong>
            </div>
          </article>
        );
      })}
    </div>
  );
};
