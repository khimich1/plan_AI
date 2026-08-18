import { useState, type CSSProperties } from "react";
import type { GsmDriver, GsmWaybill, WaybillWarningCode } from "@/features/gsm/types/gsm";
import {
  formatKm,
  formatLiters,
  formatOdometer,
  formatRouteSummary,
  isAnchorDay,
  isProblematicDay,
  warningMeta,
} from "@/features/gsm/lib/waybillWarnings";

type Props = {
  vehicleName: string;
  plateNumber: string;
  tankVolumeLiters: number;
  waybills: GsmWaybill[];
  driversById: Map<number, GsmDriver>;
  onDayClick?: (waybill: GsmWaybill) => void;
};

const stripStyle: CSSProperties = {
  display: "grid",
  gap: "0.75rem",
  padding: "1rem",
  borderRadius: 12,
  border: "1px solid #e4e7ec",
  background: "#ffffff",
};

const dayCardStyle = (anchor: boolean, problematic: boolean): CSSProperties => ({
  minWidth: 200,
  maxWidth: 240,
  flex: "0 0 auto",
  display: "grid",
  gap: "0.35rem",
  padding: "0.75rem",
  borderRadius: 10,
  border: problematic ? "1px solid #f04438" : anchor ? "1px solid #fbbf24" : "1px solid #eaecf0",
  background: problematic ? "#fef3f2" : anchor ? "#fffbeb" : "#f9fafb",
});

const badgeStyle = (active: boolean): CSSProperties => ({
  border: "none",
  borderRadius: 8,
  padding: "0.2rem 0.45rem",
  fontSize: "0.75rem",
  fontWeight: 600,
  cursor: "pointer",
  background: active ? "#fef3c7" : "#fee2e2",
  color: active ? "#92400e" : "#b42318",
});

const sparkBarStyle = (pct: number): CSSProperties => ({
  height: 6,
  width: `${Math.max(4, Math.min(100, pct))}%`,
  borderRadius: 4,
  background: "#2b5cff",
});

const WarningBadge = ({ code }: { code: WaybillWarningCode }) => {
  const [open, setOpen] = useState(false);
  const meta = warningMeta(code);
  return (
    <span style={{ display: "inline-flex", flexDirection: "column", gap: 4, alignItems: "flex-start" }}>
      <button
        type="button"
        style={badgeStyle(open)}
        aria-expanded={open}
        aria-label={`Предупреждение: ${meta.short}`}
        title={meta.reason}
        onClick={(event) => {
          event.stopPropagation();
          setOpen((v) => !v);
        }}
      >
        {meta.short}
      </button>
      {open && (
        <span role="status" style={{ fontSize: "0.8rem", color: "#475467", lineHeight: 1.35 }}>
          {meta.reason}
        </span>
      )}
    </span>
  );
};

const fuelPct = (end: number | null, tank: number): number => {
  if (end == null || !Number.isFinite(end) || tank <= 0) {
    return 0;
  }
  return (end / tank) * 100;
};

export const VehiclePeriodStrip = ({
  vehicleName,
  plateNumber,
  tankVolumeLiters,
  waybills,
  driversById,
  onDayClick,
}: Props) => {
  const sorted = [...waybills].sort((a, b) => a.date.localeCompare(b.date));
  const fuelEnds = sorted
    .map((w) => w.fuel_end)
    .filter((v): v is number => v != null && Number.isFinite(v));
  const maxFuel = Math.max(tankVolumeLiters, ...fuelEnds, 1);

  return (
    <section style={stripStyle} aria-label={`Период ${vehicleName}`}>
      <header style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem", flexWrap: "wrap" }}>
        <div>
          <h3 style={{ margin: 0, fontSize: "1.05rem" }}>
            {vehicleName}{" "}
            <span style={{ color: "#667085", fontWeight: 500 }}>({plateNumber})</span>
          </h3>
          <p style={{ margin: "0.25rem 0 0", color: "#475467", fontSize: "0.9rem" }}>
            Бак {tankVolumeLiters} л · дней {sorted.length}
            {sorted.filter(isAnchorDay).length > 0
              ? ` · якорей ${sorted.filter(isAnchorDay).length}`
              : ""}
          </p>
        </div>
      </header>

      {sorted.length === 0 ? (
        <p style={{ margin: 0, color: "#667085" }}>Нет путевых листов за выбранный период.</p>
      ) : (
        <>
          <div
            aria-label="Таймлайн остатка бака"
            style={{
              display: "flex",
              alignItems: "flex-end",
              gap: 3,
              height: 36,
              padding: "0.25rem 0",
            }}
          >
            {sorted.map((day) => (
              <div
                key={`spark-${day.id}`}
                title={`${day.date}: ${formatLiters(day.fuel_end)}`}
                style={{
                  flex: 1,
                  minWidth: 4,
                  maxWidth: 18,
                  height: `${Math.max(8, fuelPct(day.fuel_end, maxFuel))}%`,
                  borderRadius: 3,
                  background: isProblematicDay(day)
                    ? "#f04438"
                    : isAnchorDay(day)
                      ? "#f59e0b"
                      : "#93c5fd",
                }}
              />
            ))}
          </div>

          <div
            style={{
              display: "flex",
              gap: "0.65rem",
              overflowX: "auto",
              paddingBottom: "0.35rem",
            }}
          >
            {sorted.map((day) => {
              const anchor = isAnchorDay(day);
              const problematic = isProblematicDay(day);
              const driver = driversById.get(day.driver_id);
              const interactive = Boolean(onDayClick);
              return (
                <article
                  key={day.id}
                  style={{
                    ...dayCardStyle(anchor, problematic),
                    cursor: interactive ? "pointer" : undefined,
                  }}
                  data-anchor={anchor ? "true" : "false"}
                  data-problematic={problematic ? "true" : "false"}
                  aria-label={`День ${day.date}`}
                  role={interactive ? "button" : undefined}
                  tabIndex={interactive ? 0 : undefined}
                  onClick={
                    interactive
                      ? () => {
                          onDayClick?.(day);
                        }
                      : undefined
                  }
                  onKeyDown={
                    interactive
                      ? (event) => {
                          if (event.key === "Enter" || event.key === " ") {
                            event.preventDefault();
                            onDayClick?.(day);
                          }
                        }
                      : undefined
                  }
                >
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      gap: 6,
                      alignItems: "center",
                    }}
                  >
                    <strong style={{ fontSize: "0.95rem" }}>{day.date}</strong>
                    {anchor && (
                      <span
                        style={{
                          fontSize: "0.7rem",
                          fontWeight: 700,
                          color: "#92400e",
                          background: "#fde68a",
                          borderRadius: 6,
                          padding: "0.1rem 0.35rem",
                        }}
                      >
                        Якорь
                      </span>
                    )}
                  </div>
                  <div style={{ fontSize: "0.85rem", color: "#344054" }}>
                    {formatRouteSummary(day.route)}
                  </div>
                  <div style={{ fontSize: "0.85rem", color: "#475467" }}>
                    {formatKm(day.km)} · {driver?.full_name ?? `водитель #${day.driver_id}`}
                  </div>
                  <div style={{ fontSize: "0.8rem", color: "#667085" }}>
                    Топливо: {formatLiters(day.fuel_start)}
                    {(day.fuel_issued ?? 0) > 0 ? ` +${formatLiters(day.fuel_issued)}` : ""} →{" "}
                    {formatLiters(day.fuel_end)}
                  </div>
                  <div
                    aria-hidden
                    style={{
                      height: 6,
                      borderRadius: 4,
                      background: "#e4e7ec",
                      overflow: "hidden",
                    }}
                  >
                    <div style={sparkBarStyle(fuelPct(day.fuel_end, tankVolumeLiters))} />
                  </div>
                  <div style={{ fontSize: "0.8rem", color: "#667085" }}>
                    Одометр: {formatOdometer(day.odometer_start)} → {formatOdometer(day.odometer_end)}
                  </div>
                  {day.warnings.length > 0 && (
                    <div
                      style={{ display: "flex", flexWrap: "wrap", gap: 4, marginTop: 2 }}
                      aria-label="Предупреждения дня"
                    >
                      {day.warnings.map((code) => (
                        <WarningBadge key={`${day.id}-${code}`} code={code} />
                      ))}
                    </div>
                  )}
                  <div style={{ fontSize: "0.75rem", color: "#98a2b3" }}>
                    {day.status} / {day.source}
                  </div>
                </article>
              );
            })}
          </div>
        </>
      )}
    </section>
  );
};
