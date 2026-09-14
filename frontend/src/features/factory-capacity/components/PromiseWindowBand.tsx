import type { CSSProperties } from "react";
import {
  formatQuoteDayMonth,
  type PromiseQuoteWindow,
} from "@/features/factory-capacity/api/promiseQuote";

type Props = {
  window: PromiseQuoteWindow | null | undefined;
  firstPourDate?: string | null;
  pourToSunday?: string | null;
};

const wrapStyle: CSSProperties = {
  display: "grid",
  gap: "0.4rem",
  padding: "0.75rem 1rem",
  border: "1px solid #f6e05e",
  borderRadius: 12,
  background: "#fefce8",
};

const barStyle: CSSProperties = {
  position: "relative",
  height: 10,
  borderRadius: 999,
  background: "#facc15",
};

const markerStyle: CSSProperties = {
  position: "absolute",
  top: "50%",
  right: 4,
  width: 12,
  height: 12,
  borderRadius: "50%",
  background: "#92400e",
  border: "2px solid #fffbeb",
  transform: "translateY(-50%)",
  boxSizing: "border-box",
};

const metaStyle: CSSProperties = {
  display: "flex",
  justifyContent: "space-between",
  gap: "0.75rem",
  fontSize: "0.85rem",
  color: "#475467",
};

export const PromiseWindowBand = ({ window, firstPourDate, pourToSunday }: Props) => {
  if (!window) {
    return null;
  }

  const from = formatQuoteDayMonth(firstPourDate ?? window.from_week);
  const to = formatQuoteDayMonth(pourToSunday ?? window.to_week);
  const promised = formatQuoteDayMonth(window.promised_date);
  const range = from === to ? from : `${from} – ${to}`;

  return (
    <aside data-testid="promise-window-band" style={wrapStyle}>
      <div style={{ fontSize: "0.8rem", color: "#667085" }}>Окно котировки</div>
      <div style={barStyle}>
        <span
          data-testid="promise-window-band-marker"
          title={`дата клиенту ${promised}`}
          style={markerStyle}
        />
      </div>
      <div style={metaStyle}>
        <span>{range}</span>
        <strong style={{ color: "#92400e" }}>дата клиенту {promised}</strong>
      </div>
    </aside>
  );
};
