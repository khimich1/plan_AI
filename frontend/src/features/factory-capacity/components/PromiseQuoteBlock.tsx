import type { CSSProperties } from "react";
import {
  formatQuoteDayMonth,
  type PromiseQuote,
} from "@/features/factory-capacity/api/promiseQuote";

type Props = {
  quote: PromiseQuote;
};

const wrapStyle: CSSProperties = {
  display: "grid",
  gap: "0.5rem",
  padding: "0.85rem 1rem",
  border: "1px solid #e4e7ec",
  borderRadius: 12,
  background: "#fafafa",
};

const primaryStyle: CSSProperties = {
  fontSize: "1.05rem",
  fontWeight: 700,
  color: "#101828",
};

const secondaryStyle: CSSProperties = {
  display: "grid",
  gap: "0.2rem",
  fontSize: "0.85rem",
  color: "#475467",
};

function startLabel(quote: PromiseQuote): string {
  if (!quote.first_pour_date) {
    return "Начало: —";
  }
  const day = formatQuoteDayMonth(quote.first_pour_date);
  return `Начало: ${day} · остаток ${quote.first_pour_free ?? 0} дор.`;
}

export const PromiseQuoteBlock = ({ quote }: Props) => {
  const promised = formatQuoteDayMonth(quote.window?.promised_date);
  const solo = formatQuoteDayMonth(quote.solo_date);
  const soloWeekEnd = formatQuoteDayMonth(quote.solo_week_end_date);

  return (
    <aside data-testid="promise-quote-block" style={wrapStyle}>
      <div style={{ fontSize: "0.85rem", color: "#667085" }}>~{quote.tracks} дорожек</div>
      <div style={primaryStyle}>Обещать к {promised}</div>
      <div style={secondaryStyle}>
        <div>{startLabel(quote)}</div>
        <div>Если только его: {solo}</div>
        <div>Соло + до конца недели: {soloWeekEnd}</div>
      </div>
    </aside>
  );
};
