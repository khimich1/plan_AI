import type { CSSProperties } from "react";
import type { VehicleDayFeed as VehicleDayFeedData } from "@/features/gsm/lib/vehicleDayFeed";
import {
  formatKm,
  formatLiters,
  formatRouteSummary,
  warningDetailText,
  warningMeta,
} from "@/features/gsm/lib/waybillWarnings";
import type { GsmDriver, GsmTransaction, GsmWaybill } from "@/features/gsm/types/gsm";

type Props = {
  month: string; // YYYY-MM
  feed: VehicleDayFeedData[];
  driversById: Map<number, GsmDriver>;
  onGapClick?: () => void;
  onWaybillClick?: (waybill: GsmWaybill) => void;
};

const GAP_LABEL = "нет путевого на заправку/мойку";

const wrap: CSSProperties = { display: "grid", gap: "0.75rem" };

const summaryStyle: CSSProperties = {
  fontSize: "0.9rem",
  fontWeight: 600,
  color: "#344054",
};

const daySection: CSSProperties = { display: "grid", gap: "0.4rem" };

const dayTitle: CSSProperties = {
  margin: 0,
  fontSize: "0.85rem",
  fontWeight: 600,
  color: "#475467",
};

const cardBase: CSSProperties = {
  border: "1px solid #eaecf0",
  borderRadius: 10,
  background: "#ffffff",
  padding: "0.5rem 0.75rem",
  fontSize: "0.85rem",
  color: "#344054",
  display: "flex",
  gap: "0.75rem",
  flexWrap: "wrap",
  alignItems: "baseline",
};

const gapCard: CSSProperties = {
  ...cardBase,
  border: "1px solid #fdb022",
  background: "#fffaeb",
  alignItems: "center",
};

const gapText: CSSProperties = { color: "#93370d", fontSize: "0.8rem" };

const gapCta: CSSProperties = {
  marginLeft: "auto",
  border: "1px solid #fdb022",
  borderRadius: 8,
  background: "#ffffff",
  color: "#93370d",
  padding: "0.25rem 0.6rem",
  fontSize: "0.8rem",
  fontWeight: 600,
  cursor: "pointer",
};

const wbCard: CSSProperties = {
  ...cardBase,
  cursor: "pointer",
  textAlign: "left",
  width: "100%",
  boxSizing: "border-box",
  font: "inherit",
};

const muted: CSSProperties = { color: "#667085" };

const badge = (code: string): CSSProperties => ({
  display: "inline-block",
  marginRight: 4,
  fontSize: "0.75rem",
  background: code === "manual_intervention" ? "#fee4e2" : "#fef0c7",
  color: code === "manual_intervention" ? "#b42318" : "#93370d",
  borderRadius: 8,
  padding: "0.1rem 0.4rem",
});

const formatAmount = (value: number): string =>
  `${value.toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽`;

const serviceLabel = (type: string): string => {
  if (type === "fuel") return "Топливо";
  if (type === "wash") return "Мойка";
  return type;
};

const dayTitleText = (iso: string): string =>
  new Date(`${iso}T12:00:00`).toLocaleDateString("ru-RU", {
    weekday: "short",
    day: "numeric",
    month: "long",
  });

const monthTitleText = (month: string): string =>
  new Date(`${month}-01T12:00:00`).toLocaleDateString("ru-RU", {
    month: "long",
    year: "numeric",
  });

const TxCard = ({ tx }: { tx: GsmTransaction }) => (
  <div style={cardBase} data-testid={`feed-tx-${tx.ts}`}>
    <span>{tx.ts.slice(11, 16)}</span>
    <span>{serviceLabel(tx.service_type)}</span>
    <span style={muted}>{tx.address ?? "—"}</span>
    <span>{formatLiters(tx.qty_liters)}</span>
    <span>{formatAmount(tx.amount)}</span>
  </div>
);

export const VehicleDayFeed = ({
  month,
  feed,
  driversById,
  onGapClick,
  onWaybillClick,
}: Props) => {
  if (feed.length === 0) return null;

  const waybills = feed.flatMap((day) => day.waybills);
  const sumKm = waybills.reduce((sum, wb) => sum + (wb.km || 0), 0);
  const sumIssued = waybills.reduce((sum, wb) => sum + (wb.fuel_issued || 0), 0);

  return (
    <div style={wrap} data-testid="vehicle-day-feed">
      <div style={summaryStyle} data-testid="feed-summary">
        Итого за {monthTitleText(month)}: {waybills.length} ПЛ, {formatKm(sumKm)}, выдано{" "}
        {formatLiters(sumIssued)}
      </div>
      {feed.map((day) => (
        <section key={day.date} style={daySection} data-testid={`feed-day-${day.date}`}>
          <h3 style={dayTitle}>{dayTitleText(day.date)}</h3>
          {day.txs.map((tx, index) =>
            day.isGap ? (
              <div
                key={`${tx.ts}-${index}`}
                role="region"
                aria-label={`${GAP_LABEL}: ${day.date}`}
                data-testid={`feed-gap-${day.date}`}
                style={gapCard}
              >
                <span>{tx.ts.slice(11, 16)}</span>
                <span>{serviceLabel(tx.service_type)}</span>
                <span style={muted}>{tx.address ?? "—"}</span>
                <span>{formatLiters(tx.qty_liters)}</span>
                <span>{formatAmount(tx.amount)}</span>
                <span style={gapText}>{GAP_LABEL}</span>
                {onGapClick && (
                  <button type="button" style={gapCta} onClick={onGapClick}>
                    Сгенерировать
                  </button>
                )}
              </div>
            ) : (
              <TxCard key={`${tx.ts}-${index}`} tx={tx} />
            ),
          )}
          {day.waybills.map((wb) => (
            <button
              key={wb.id}
              type="button"
              style={wbCard}
              data-testid={`feed-wb-${wb.id}`}
              onClick={() => onWaybillClick?.(wb)}
            >
              <span>{driversById.get(wb.driver_id)?.full_name ?? wb.driver_id}</span>
              <span style={muted}>{formatRouteSummary(wb.route)}</span>
              <span>{formatKm(wb.km)}</span>
              <span>
                {formatLiters(wb.fuel_start)} / {formatLiters(wb.fuel_issued)} /{" "}
                {formatLiters(wb.fuel_end)}
              </span>
              <span>{wb.status}</span>
              <span>
                {wb.warnings.map((code) => (
                  <span key={code} title={warningDetailText(wb, code)} style={badge(code)}>
                    {warningMeta(code).short}
                    {wb.warning_details?.find((d) => d.code === code)?.detail
                      ? `: ${wb.warning_details.find((d) => d.code === code)?.detail}`
                      : ""}
                  </span>
                ))}
              </span>
            </button>
          ))}
        </section>
      ))}
    </div>
  );
};
