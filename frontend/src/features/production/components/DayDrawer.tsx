import { useMemo } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Drawer } from "@/shared/ui/Drawer";
import { Spinner } from "@/shared/ui/Spinner";
import {
  useCompleteDayMutation,
  useDayDocumentMutation,
  useDayViewQuery,
} from "@/features/production/hooks/useProductionQueries";
import type {
  DayInfo,
  DayPlanBlock,
  DayTrackDetail,
} from "@/features/production/types/production";

type DayDrawerProps = {
  date: string | null;
  summary?: DayInfo;
  onClose: () => void;
};

const formatDateRu = (iso: string): string => {
  const [year, month, day] = iso.split("-").map(Number);
  if (!year || !month || !day) return iso;
  const d = new Date(Date.UTC(year, month - 1, day));
  return d.toLocaleDateString("ru-RU", {
    day: "2-digit",
    month: "long",
    year: "numeric",
    timeZone: "UTC",
  });
};

const formatSize = (track: DayTrackDetail): string => {
  const parts: string[] = [];
  if (typeof track.length === "number" && track.length > 0) {
    parts.push(`L=${track.length.toFixed(2)} м`);
  }
  if (track.max_reinforcement > 0) {
    parts.push(`арм. ${track.max_reinforcement.toFixed(1)}`);
  }
  if (track.label) {
    parts.push(track.label);
  }
  return parts.join(" · ");
};

const formatLengthM = (value: number): string => {
  if (Math.abs(value - Math.round(value)) < 0.005) {
    return `${Math.round(value)}`;
  }
  return value.toFixed(2).replace(/0+$/, "").replace(/\.$/, "");
};

export const DayDrawer = ({ date, summary, onClose }: DayDrawerProps) => {
  const open = date !== null;
  const dayQuery = useDayViewQuery(date);
  const completeMutation = useCompleteDayMutation();
  const schemaMutation = useDayDocumentMutation("schema");
  const breakdownMutation = useDayDocumentMutation("breakdown");
  const formovkaMutation = useDayDocumentMutation("formovka");

  const plans: DayPlanBlock[] = useMemo(
    () => dayQuery.data?.plans ?? [],
    [dayQuery.data],
  );

  const totalTracks = dayQuery.data?.total_tracks ?? 0;
  const hasTracks = totalTracks > 0;

  const handleCompleteDay = (planId: string) => {
    if (date) {
      completeMutation.mutate({ date, planId });
    }
  };

  const anyDocumentLoading =
    schemaMutation.isPending ||
    breakdownMutation.isPending ||
    formovkaMutation.isPending;

  return (
    <Drawer
      open={open}
      onClose={onClose}
      width={860}
      title={
        date ? (
          <div style={{ display: "grid", gap: "0.15rem" }}>
            <span>{formatDateRu(date)}</span>
            {summary && (
              <span style={{ fontSize: "0.85rem", color: "#475467", fontWeight: 400 }}>
                Занято дорожек: <strong>{summary.occupied}</strong>/{summary.max}
                {summary.completed && (
                  <span style={{ marginLeft: 8, color: "#067647" }}>✓ Выполнен</span>
                )}
              </span>
            )}
          </div>
        ) : null
      }
    >
      {dayQuery.isLoading && (
        <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
          <Spinner /> Загрузка содержимого дня…
        </div>
      )}

      {dayQuery.isError && (
        <Alert tone="error">Не удалось загрузить информацию о дне.</Alert>
      )}

      {dayQuery.data && !hasTracks && (
        <Alert tone="info">На этот день не запланировано дорожек.</Alert>
      )}

      {dayQuery.data && hasTracks && date && (
        <div style={{ display: "grid", gap: "1rem" }}>
          <section>
            <h3 style={{ margin: "0 0 0.5rem", fontSize: "1rem" }}>Документы</h3>
            <div className="day-docs">
              <Button
                variant="primary"
                onClick={() => schemaMutation.mutate(date)}
                disabled={anyDocumentLoading}
              >
                {schemaMutation.isPending ? "Формируем…" : "Схема (PDF)"}
              </Button>
              <Button
                variant="secondary"
                onClick={() => breakdownMutation.mutate(date)}
                disabled={anyDocumentLoading}
              >
                {breakdownMutation.isPending ? "Формируем…" : "Разбивка (XLSX)"}
              </Button>
              <Button
                variant="secondary"
                onClick={() => formovkaMutation.mutate(date)}
                disabled={anyDocumentLoading}
              >
                {formovkaMutation.isPending ? "Формируем…" : "Формовка (ZIP)"}
              </Button>
            </div>
            {(schemaMutation.isError ||
              breakdownMutation.isError ||
              formovkaMutation.isError) && (
              <div style={{ marginTop: "0.5rem" }}>
                <Alert tone="error">
                  Не удалось сформировать документ. Попробуйте ещё раз.
                </Alert>
              </div>
            )}
          </section>

          {plans.map((plan) => (
            <div className="day-plan-block" key={plan.plan_id}>
              <header className="day-plan-block__header">
                <div className="day-plan-block__header-title">
                  {plan.plan_name || plan.plan_id}
                  <span
                    style={{
                      marginLeft: 8,
                      color: "#475467",
                      fontWeight: 400,
                      fontSize: "0.85rem",
                    }}
                  >
                    · {plan.tracks.length} дор.
                  </span>
                  {plan.completed && (
                    <span className="day-plan-block__completed" style={{ marginLeft: 10 }}>
                      ✓ Выполнен
                    </span>
                  )}
                </div>
                <Button
                  variant="secondary"
                  onClick={() => handleCompleteDay(plan.plan_id)}
                  disabled={completeMutation.isPending || plan.completed}
                >
                  {plan.completed
                    ? "Уже отмечен"
                    : completeMutation.isPending
                      ? "Сохранение…"
                      : "Отметить выполненным"}
                </Button>
              </header>

              {plan.tracks.map((track) => (
                <div className="day-track" key={`${plan.plan_id}-${track.track_number}`}>
                  <div className="day-track__header">
                    <div className="day-track__title">Дорожка {track.track_number}</div>
                    <div className="day-track__meta">{formatSize(track)}</div>
                  </div>

                  {track.plates_info.length === 0 ? (
                    <div style={{ color: "#98a2b3", fontSize: "0.9rem" }}>
                      Плит не найдено
                    </div>
                  ) : (
                    <table className="day-plates-table">
                      <thead>
                        <tr>
                          <th>Плита</th>
                          <th>Размер</th>
                          <th>Заказчик</th>
                          <th>Срок КП</th>
                          <th className="day-plates-table__qty">Кол-во</th>
                        </tr>
                      </thead>
                      <tbody>
                        {track.plates_info.map((plate, index) => (
                          <tr key={`${track.track_number}-${index}`}>
                            <td>
                              <strong>{plate.plate_name || "—"}</strong>
                              {plate.kp_id && (
                                <span
                                  style={{
                                    color: "#98a2b3",
                                    marginLeft: 6,
                                    fontSize: "0.8rem",
                                  }}
                                >
                                  КП {plate.kp_id}
                                </span>
                              )}
                            </td>
                            <td>
                              {formatLengthM(plate.length_m)} м ×{" "}
                              {plate.width_mm} мм
                            </td>
                            <td>{plate.customer}</td>
                            <td>{plate.kp_date}</td>
                            <td className="day-plates-table__qty">
                              {plate.qty}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              ))}
            </div>
          ))}

          {completeMutation.isError && (
            <Alert tone="error">Не удалось отметить день выполненным.</Alert>
          )}
        </div>
      )}
    </Drawer>
  );
};
