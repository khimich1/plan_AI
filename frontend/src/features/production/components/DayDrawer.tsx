import { useEffect, useMemo, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Drawer } from "@/shared/ui/Drawer";
import { Spinner } from "@/shared/ui/Spinner";
import { getErrorMessage } from "@/shared/lib/apiError";
import { isPlanVersionConflict } from "@/shared/lib/planConflict";
import {
  useCompleteDayMutation,
  useDayDocumentMutation,
  useDayViewQuery,
  useDeleteTrackMutation,
  usePlansListQuery,
} from "@/features/production/hooks/useProductionQueries";
import type { BasketDayKind } from "@/features/production/lib/basketDayKind";
import type {
  DayInfo,
  DayPlanBlock,
  DayTrackDetail,
  RejectedPlateItem,
} from "@/features/production/types/production";

type DayDrawerProps = {
  date: string | null;
  summary?: DayInfo;
  /** Kind текущей корзины — зарезервирован для будущих подсказок. */
  basketKind?: BasketDayKind | null;
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

const makeRejectionKey = (
  planId: string,
  trackNumber: number,
  plateIndex: number,
): string => `${planId}:${trackNumber}:${plateIndex}`;

const countTrackPlates = (track: DayTrackDetail): number =>
  track.plates_info.reduce((sum, plate) => sum + plate.qty, 0);

export const DayDrawer = ({
  date,
  summary,
  onClose,
}: DayDrawerProps) => {
  const open = date !== null;
  const dayQuery = useDayViewQuery(date);
  const plansQuery = usePlansListQuery();
  const completeMutation = useCompleteDayMutation();
  const deleteTrackMutation = useDeleteTrackMutation();
  const schemaMutation = useDayDocumentMutation("schema");
  const breakdownMutation = useDayDocumentMutation("breakdown");
  const formovkaMutation = useDayDocumentMutation("formovka");
  const [rejectedByPlate, setRejectedByPlate] = useState<Record<string, number>>({});

  useEffect(() => {
    setRejectedByPlate({});
  }, [date]);

  const plans: DayPlanBlock[] = useMemo(
    () => dayQuery.data?.plans ?? [],
    [dayQuery.data],
  );

  const planVersionById = useMemo(() => {
    const versions = new Map<string, number>();
    for (const plan of plansQuery.data?.plans ?? []) {
      if (typeof plan.version === "number") {
        versions.set(plan.id, plan.version);
      }
    }
    return versions;
  }, [plansQuery.data]);

  const getExpectedVersion = (planId: string): number | undefined =>
    planVersionById.get(planId);

  const totalTracks = dayQuery.data?.total_tracks ?? 0;
  const hasTracks = totalTracks > 0;

  const setRejectedQty = (
    planId: string,
    trackNumber: number,
    plateIndex: number,
    nextQty: number,
    maxQty: number,
  ) => {
    const key = makeRejectionKey(planId, trackNumber, plateIndex);
    const clampedQty = Math.max(0, Math.min(maxQty, nextQty));

    setRejectedByPlate((current) => {
      const updated = { ...current };
      if (clampedQty === 0) {
        delete updated[key];
      } else {
        updated[key] = clampedQty;
      }
      return updated;
    });
  };

  const clearPlanRejections = (planId: string) => {
    setRejectedByPlate((current) => {
      const updated = { ...current };
      for (const key of Object.keys(updated)) {
        if (key.startsWith(`${planId}:`)) {
          delete updated[key];
        }
      }
      return updated;
    });
  };

  const buildRejectedPlates = (plan: DayPlanBlock): RejectedPlateItem[] =>
    plan.tracks.flatMap((track) =>
      track.plates_info.flatMap((plate, plateIndex) => {
        const key = makeRejectionKey(plan.plan_id, track.track_number, plateIndex);
        const qty = rejectedByPlate[key] ?? 0;
        if (qty <= 0) {
          return [];
        }
        return [
          {
            track_number: track.track_number,
            plate_index: plateIndex,
            qty: Math.min(qty, plate.qty),
          },
        ];
      }),
    );

  const handleCompleteDay = (plan: DayPlanBlock) => {
    if (date) {
      completeMutation.mutate(
        {
          date,
          planId: plan.plan_id,
          rejectedPlates: buildRejectedPlates(plan),
          expectedVersion: getExpectedVersion(plan.plan_id),
        },
        {
          onSuccess: () => clearPlanRejections(plan.plan_id),
          onError: (error) => {
            if (isPlanVersionConflict(error)) {
              clearPlanRejections(plan.plan_id);
            }
          },
        },
      );
    }
  };

  const handleDeleteTrack = (plan: DayPlanBlock, track: DayTrackDetail) => {
    if (!date || plan.completed) return;

    const platesCount = countTrackPlates(track);
    const platesHint =
      platesCount > 0
        ? `${platesCount} шт. плит вернётся в производство.`
        : "Плиты с дорожки вернутся в производство.";
    const confirmed = window.confirm(
      `Удалить дорожку ${track.track_number}? ${platesHint}`,
    );
    if (!confirmed) return;

    deleteTrackMutation.mutate({
      planId: plan.plan_id,
      date,
      trackIndex: track.plan_track_index,
      expectedVersion: getExpectedVersion(plan.plan_id),
    });
  };

  const isDeletingTrack = (planId: string, trackIndex: number): boolean =>
    deleteTrackMutation.isPending &&
    deleteTrackMutation.variables?.planId === planId &&
    deleteTrackMutation.variables?.trackIndex === trackIndex;

  const anyDocumentLoading =
    schemaMutation.isPending ||
    breakdownMutation.isPending ||
    formovkaMutation.isPending;
  const completionResult =
    completeMutation.data?.date === date ? completeMutation.data : null;
  const deleteTrackResult =
    deleteTrackMutation.data?.date === date ? deleteTrackMutation.data : null;
  const completeErrorMessage =
    completeMutation.isError && !isPlanVersionConflict(completeMutation.error)
      ? getErrorMessage(completeMutation.error)
      : null;
  const deleteTrackErrorMessage = deleteTrackMutation.isError
    ? (() => {
        const error = deleteTrackMutation.error;
        if (isPlanVersionConflict(error)) {
          return null;
        }
        if (error.status === 409) {
          return "День уже завершён, удаление невозможно";
        }
        return getErrorMessage(error);
      })()
    : null;

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

          {completionResult && (
            <Alert tone="success">
              День отмечен выполненным. Списано:{" "}
              {completionResult.moved_plates ?? 0}, возвращено брака:{" "}
              {completionResult.rejected_returned ?? 0}.
            </Alert>
          )}

          {deleteTrackResult && (
            <Alert tone="success">
              Дорожка удалена. В производство возвращено:{" "}
              {deleteTrackResult.plates_returned} плит.
            </Alert>
          )}

          {deleteTrackErrorMessage && (
            <Alert tone="error">{deleteTrackErrorMessage}</Alert>
          )}

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
                  onClick={() => handleCompleteDay(plan)}
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
                  <div
                    className="day-track__header"
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "flex-start",
                      gap: "0.75rem",
                    }}
                  >
                    <div>
                      <div className="day-track__title">Дорожка {track.track_number}</div>
                      <div className="day-track__meta">{formatSize(track)}</div>
                    </div>
                    {!plan.completed && (
                      <Button
                        variant="secondary"
                        onClick={() => handleDeleteTrack(plan, track)}
                        disabled={
                          deleteTrackMutation.isPending ||
                          completeMutation.isPending ||
                          plan.completed
                        }
                      >
                        {isDeletingTrack(plan.plan_id, track.plan_track_index)
                          ? "Удаление…"
                          : "Удалить дорожку"}
                      </Button>
                    )}
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
                          <th className="day-plates-table__qty">Брак</th>
                          <th className="day-plates-table__qty">Выполнено</th>
                        </tr>
                      </thead>
                      <tbody>
                        {track.plates_info.map((plate, index) => {
                          const key = makeRejectionKey(
                            plan.plan_id,
                            track.track_number,
                            index,
                          );
                          const rejectedQty = rejectedByPlate[key] ?? 0;
                          const completedQty = Math.max(plate.qty - rejectedQty, 0);
                          const plateWrittenOff = Boolean(plate.write_off_completed);
                          const controlsDisabled =
                            plan.completed ||
                            completeMutation.isPending ||
                            plateWrittenOff;

                          return (
                            <tr
                              key={`${track.track_number}-${index}`}
                              className={
                                plateWrittenOff ? "day-plates-table__row--written-off" : undefined
                              }
                            >
                              <td>
                                <div style={{ display: "flex", alignItems: "center", gap: "0.5rem", flexWrap: "wrap" }}>
                                  <strong>{plate.plate_name || "—"}</strong>
                                  {plateWrittenOff && (
                                    <span className="day-plate-badge day-plate-badge--done">
                                      (ГОТОВО)
                                    </span>
                                  )}
                                </div>
                                {plate.kp_id ? (
                                  <div style={{ color: "#98a2b3", marginTop: 2, fontSize: "0.8rem" }}>
                                    КП {plate.kp_id}
                                  </div>
                                ) : null}
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
                              <td className="day-plates-table__qty">
                                <div className="day-reject-control">
                                  <button
                                    type="button"
                                    onClick={() =>
                                      setRejectedQty(
                                        plan.plan_id,
                                        track.track_number,
                                        index,
                                        rejectedQty - 1,
                                        plate.qty,
                                      )
                                    }
                                    disabled={controlsDisabled || rejectedQty <= 0}
                                  >
                                    -
                                  </button>
                                  <span>{rejectedQty}</span>
                                  <button
                                    type="button"
                                    onClick={() =>
                                      setRejectedQty(
                                        plan.plan_id,
                                        track.track_number,
                                        index,
                                        rejectedQty + 1,
                                        plate.qty,
                                      )
                                    }
                                    disabled={controlsDisabled || rejectedQty >= plate.qty}
                                  >
                                    +
                                  </button>
                                  {rejectedQty > 0 && (
                                    <button
                                      type="button"
                                      onClick={() =>
                                        setRejectedQty(
                                          plan.plan_id,
                                          track.track_number,
                                          index,
                                          0,
                                          plate.qty,
                                        )
                                      }
                                      disabled={controlsDisabled}
                                    >
                                      Сброс
                                    </button>
                                  )}
                                </div>
                              </td>
                              <td className="day-plates-table__qty">
                                {completedQty}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  )}
                </div>
              ))}
            </div>
          ))}

          {completeErrorMessage && (
            <Alert tone="error">
              Не удалось отметить день выполненным: {completeErrorMessage}
            </Alert>
          )}
        </div>
      )}
    </Drawer>
  );
};
