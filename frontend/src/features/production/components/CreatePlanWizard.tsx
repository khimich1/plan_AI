import { useEffect, useMemo, useRef, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import { MonthCalendarGrid } from "@/features/production/components/MonthCalendarGrid";
import {
  useBuildPlanMutation,
  useDayOccupancyQuery,
  useGlobalCalendarQuery,
  useKpCandidatesQuery,
  useWorkCalendarQuery,
} from "@/features/production/hooks/useProductionQueries";
import type {
  FillTargetItem,
  FilterMethod,
  KpCandidateItem,
} from "@/features/production/types/production";

type Props = {
  onCreated?: () => void;
  /** Если задано — мастер открывается сразу на шаге 3 в режиме «дозаполнение». */
  fillRequest?: FillTargetItem[] | null;
  /** Сигнализирует родителю, что fillRequest подхвачен и можно его очистить. */
  onFillRequestConsumed?: () => void;
  /** Возврат к календарю при отмене дозаполнения. */
  onCancelFill?: () => void;
};

const PRESET_TRACKS = [1, 2, 3, 4, 5];
const MAX_PER_DAY = 5;

const todayISO = () => {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
};

const parseISODate = (iso: string) => {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
};

const startOfMonth = (d: Date) => new Date(d.getFullYear(), d.getMonth(), 1);

const formatRu = (iso: string) => {
  const d = parseISODate(iso);
  return d.toLocaleDateString("ru-RU", { day: "2-digit", month: "2-digit" });
};

export const CreatePlanWizard = ({
  onCreated,
  fillRequest,
  onFillRequestConsumed,
  onCancelFill,
}: Props) => {
  const [step, setStep] = useState<1 | 2 | 3>(1);
  const [startDate, setStartDate] = useState<string>(todayISO());
  const [tracksCount, setTracksCount] = useState<number>(1);
  const [filterMethod, setFilterMethod] = useState<FilterMethod>("all");
  const [selectedPlatesByKp, setSelectedPlatesByKp] = useState<
    Record<number, number[]>
  >({});
  const [expandedKpIds, setExpandedKpIds] = useState<Set<number>>(new Set());
  const [planName, setPlanName] = useState<string>("");
  const [calendarMonth, setCalendarMonth] = useState<Date>(() =>
    startOfMonth(new Date()),
  );

  // Режим «дозаполнение дней»: список целей (даты + сколько дорожек). Когда
  // он не null — индикатор шагов скрыт, шаг 1/2 пропускаются, в submit
  // уходит fill_targets, и активный план не дописывается.
  const [fillTargets, setFillTargets] = useState<FillTargetItem[] | null>(null);

  const occupancyQuery = useDayOccupancyQuery();
  const calendarQuery = useGlobalCalendarQuery();
  const workCalendar = useWorkCalendarQuery();
  const candidatesQuery = useKpCandidatesQuery(step === 3 && filterMethod === "kp");
  const buildMutation = useBuildPlanMutation();

  const occupancy = occupancyQuery.data?.occupancy ?? {};
  const maxPerDay = occupancyQuery.data?.max_per_day ?? MAX_PER_DAY;
  const daysInfo = calendarQuery.data?.days_info ?? {};
  const holidays = useMemo(
    () => new Set(workCalendar.data?.extra_holidays ?? []),
    [workCalendar.data],
  );
  const extraWorkdays = useMemo(
    () => new Set(workCalendar.data?.extra_workdays ?? []),
    [workCalendar.data],
  );

  const selectedDay = startDate;
  const occupiedOnStart = occupancy[selectedDay] ?? 0;
  const freeOnStart = Math.max(0, maxPerDay - occupiedOnStart);

  // Когда родитель передал корзину дней — переходим сразу на шаг 3.
  useEffect(() => {
    if (fillRequest && fillRequest.length > 0) {
      setFillTargets(fillRequest);
      setStep(3);
      onFillRequestConsumed?.();
    }
  }, [fillRequest, onFillRequestConsumed]);

  const toggleKp = (kp: KpCandidateItem) => {
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      if (kp.kp_id in next) {
        delete next[kp.kp_id];
      } else {
        next[kp.kp_id] = kp.plates.map((p) => p.id);
      }
      return next;
    });
  };

  const togglePlate = (kp: KpCandidateItem, plateId: number) => {
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      const current = next[kp.kp_id];
      if (current === undefined) {
        next[kp.kp_id] = [plateId];
      } else if (current.includes(plateId)) {
        const filtered = current.filter((id) => id !== plateId);
        if (filtered.length === 0) {
          delete next[kp.kp_id];
        } else {
          next[kp.kp_id] = filtered;
        }
      } else {
        next[kp.kp_id] = [...current, plateId];
      }
      return next;
    });
  };

  const toggleExpand = (kpId: number) => {
    setExpandedKpIds((prev) => {
      const next = new Set(prev);
      if (next.has(kpId)) {
        next.delete(kpId);
      } else {
        next.add(kpId);
      }
      return next;
    });
  };

  const isFillMode = fillTargets !== null;
  const canProceedStep1 = Boolean(startDate);
  const canProceedStep2 = tracksCount >= 1 && tracksCount <= 50;
  const hasAnyPlateSelected =
    filterMethod === "all" ||
    Object.values(selectedPlatesByKp).some((ids) => ids.length > 0);
  const canSubmit =
    (isFillMode || (canProceedStep1 && canProceedStep2)) &&
    hasAnyPlateSelected &&
    !buildMutation.isPending &&
    !buildMutation.isSuccess;

  const fillTotalTracks = fillTargets
    ? fillTargets.reduce((acc, t) => acc + t.tracks, 0)
    : 0;
  const fillSubtitle = fillTargets
    ? `${fillTotalTracks} дор. на ${fillTargets.length} ` +
      `${fillTargets.length === 1 ? "день" : "днях"}: ` +
      fillTargets.map((t) => `${formatRu(t.date)} (${t.tracks})`).join(", ") +
      ". Лишние плиты остаются «в производстве»."
    : undefined;

  const handleSubmit = () => {
    const selectedKpIds = Object.entries(selectedPlatesByKp)
      .filter(([, ids]) => ids.length > 0)
      .map(([kpId]) => Number(kpId));

    let partialPlateIds: Record<number, number[]> | undefined;
    if (filterMethod === "kp" && candidatesQuery.data) {
      const candidatesByKp = new Map(
        candidatesQuery.data.items.map((kp) => [kp.kp_id, kp]),
      );
      const partial: Record<number, number[]> = {};
      for (const [kpIdStr, plateIds] of Object.entries(selectedPlatesByKp)) {
        const kpId = Number(kpIdStr);
        const kp = candidatesByKp.get(kpId);
        if (!kp || plateIds.length === 0) {
          continue;
        }
        if (plateIds.length < kp.plates.length) {
          partial[kpId] = plateIds;
        }
      }
      if (Object.keys(partial).length > 0) {
        partialPlateIds = partial;
      }
    }

    // В режиме fill_targets бэк сам пересчитает start_date / tracks_count из
    // массива таргетов. Передаём здесь min(date) и max(tracks), чтобы запрос
    // прошёл базовую Pydantic-валидацию (start_date != "", tracks_count >= 1).
    const fillStart = fillTargets ? fillTargets[0].date : startDate;
    const fillTracks = fillTargets
      ? Math.max(...fillTargets.map((t) => t.tracks))
      : tracksCount;

    buildMutation.mutate(
      {
        start_date: fillStart,
        tracks_count: fillTracks,
        filter_method: filterMethod,
        selected_kp_ids: filterMethod === "kp" ? selectedKpIds : undefined,
        selected_plate_ids: partialPlateIds,
        plan_name: planName.trim() ? planName.trim() : undefined,
        fill_targets: fillTargets ?? undefined,
      },
      {
        onSuccess: () => {
          setStep(1);
          setSelectedPlatesByKp({});
          setExpandedKpIds(new Set());
          setPlanName("");
          setFillTargets(null);
          onCreated?.();
        },
      },
    );
  };

  const handleCancelFill = () => {
    setFillTargets(null);
    setStep(1);
    setSelectedPlatesByKp({});
    setExpandedKpIds(new Set());
    onCancelFill?.();
  };

  const cardTitle = isFillMode ? "Дозаполнение дней" : "Начать планирование";
  const cardSubtitle = isFillMode
    ? fillSubtitle
    : "Мастер создания нового производственного плана в три шага.";

  return (
    <Card title={cardTitle} subtitle={cardSubtitle}>
      {!isFillMode && (
        <div style={{ display: "flex", gap: "0.5rem", marginBottom: "1rem" }}>
          {[1, 2, 3].map((s) => (
            <div
              key={s}
              style={{
                flex: 1,
                padding: "0.5rem 0.75rem",
                borderRadius: 10,
                background: s === step ? "#2b5cff" : "#eef2ff",
                color: s === step ? "#ffffff" : "#23366f",
                fontWeight: 600,
                textAlign: "center",
              }}
            >
              Шаг {s}
            </div>
          ))}
        </div>
      )}

      {!isFillMode && step === 1 && (
        <div style={{ display: "grid", gap: "1rem" }}>
          <FieldWrapper
            label="Дата начала планирования"
            hint="План построится начиная с этого дня (рабочие дни учитываются автоматически)."
          >
            <Input
              type="date"
              value={startDate}
              onChange={(e) => setStartDate(e.target.value)}
            />
          </FieldWrapper>

          <FieldWrapper label="Название плана (необязательно)">
            <Input
              type="text"
              placeholder="Например: «План на 20.04»"
              value={planName}
              onChange={(e) => setPlanName(e.target.value)}
            />
          </FieldWrapper>

          <div>
            <div style={{ fontWeight: 600, marginBottom: "0.5rem" }}>
              Загрузка дат (из других планов):
            </div>
            {calendarQuery.isLoading || workCalendar.isLoading ? (
              <Spinner />
            ) : (
              <MonthCalendarGrid
                daysInfo={daysInfo}
                holidays={holidays}
                extraWorkdays={extraWorkdays}
                month={calendarMonth}
                onMonthChange={setCalendarMonth}
                selectedDate={startDate}
                onSelectDate={setStartDate}
              />
            )}
            <div style={{ color: "#475467", marginTop: "0.35rem" }}>
              На выбранный день уже занято <strong>{occupiedOnStart}</strong> из{" "}
              {maxPerDay} дорожек (свободно {freeOnStart}).
            </div>
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end" }}>
            <Button onClick={() => setStep(2)} disabled={!canProceedStep1}>
              Далее →
            </Button>
          </div>
        </div>
      )}

      {!isFillMode && step === 2 && (
        <div style={{ display: "grid", gap: "1rem" }}>
          <FieldWrapper
            label="Количество дорожек в день"
            hint="От 1 до 50. Максимум одновременно задействованных дорожек на одной дате."
          >
            <Input
              type="number"
              min={1}
              max={50}
              value={tracksCount}
              onChange={(e) =>
                setTracksCount(Math.max(1, Math.min(50, Number(e.target.value) || 1)))
              }
            />
          </FieldWrapper>

          <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
            {PRESET_TRACKS.map((preset) => (
              <Button
                key={preset}
                variant={tracksCount === preset ? "primary" : "secondary"}
                onClick={() => setTracksCount(preset)}
              >
                {preset}
              </Button>
            ))}
          </div>

          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <Button variant="ghost" onClick={() => setStep(1)}>
              ← Назад
            </Button>
            <Button onClick={() => setStep(3)} disabled={!canProceedStep2}>
              Далее →
            </Button>
          </div>
        </div>
      )}

      {step === 3 && (
        <div style={{ display: "grid", gap: "1rem" }}>
          {isFillMode && (
            <FieldWrapper label="Название плана (необязательно)">
              <Input
                type="text"
                placeholder="Например: «Дозаполнение 27-28.04»"
                value={planName}
                onChange={(e) => setPlanName(e.target.value)}
              />
            </FieldWrapper>
          )}

          <FieldWrapper label="Какие КП включить в план">
            <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
              <Button
                variant={filterMethod === "all" ? "primary" : "secondary"}
                onClick={() => setFilterMethod("all")}
              >
                Все КП в работе
              </Button>
              <Button
                variant={filterMethod === "kp" ? "primary" : "secondary"}
                onClick={() => setFilterMethod("kp")}
              >
                Выбрать КП вручную
              </Button>
            </div>
          </FieldWrapper>

          {filterMethod === "kp" && (
            <div style={{ border: "1px solid #e4e7ec", borderRadius: 14, overflow: "hidden" }}>
              {candidatesQuery.isLoading && (
                <div style={{ padding: "1rem", display: "flex", gap: "0.5rem" }}>
                  <Spinner /> Загрузка КП…
                </div>
              )}
              {candidatesQuery.isError && (
                <div style={{ padding: "1rem" }}>
                  <Alert tone="error">Не удалось загрузить список КП.</Alert>
                </div>
              )}
              {candidatesQuery.data && (
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead style={{ background: "#f8fafc" }}>
                    <tr>
                      <th style={thStyle}></th>
                      <th style={thStyle}>Выбор</th>
                      <th style={thStyle}>КП №</th>
                      <th style={thStyle}>Заказчик</th>
                      <th style={thStyle}>Срок</th>
                      <th style={thStyle}>Выполнено</th>
                      <th style={thStyle}>В плане</th>
                      <th style={thStyle}>Длина, м</th>
                      <th style={thStyle}>Плиты</th>
                    </tr>
                  </thead>
                  <tbody>
                    {candidatesQuery.data.items.length === 0 && (
                      <tr>
                        <td
                          colSpan={9}
                          style={{ padding: "1rem", textAlign: "center", color: "#475467" }}
                        >
                          Нет КП в работе с неразмещёнными плитами.
                        </td>
                      </tr>
                    )}
                    {candidatesQuery.data.items.map((kp) => {
                      const totalPlates = kp.plates.length;
                      const selected = selectedPlatesByKp[kp.kp_id];
                      const selectedCount = selected?.length ?? 0;
                      const isChecked = selectedCount === totalPlates && totalPlates > 0;
                      const isIndeterminate =
                        selectedCount > 0 && selectedCount < totalPlates;
                      const isExpanded = expandedKpIds.has(kp.kp_id);
                      return (
                        <ExpandableKpRow
                          key={kp.kp_id}
                          kp={kp}
                          isExpanded={isExpanded}
                          isChecked={isChecked}
                          isIndeterminate={isIndeterminate}
                          selectedCount={selectedCount}
                          selectedIds={selected ?? []}
                          onToggleKp={() => toggleKp(kp)}
                          onToggleExpand={() => toggleExpand(kp.kp_id)}
                          onTogglePlate={(plateId) => togglePlate(kp, plateId)}
                        />
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>
          )}

          {buildMutation.isError && (
            <Alert tone="error">
              {(buildMutation.error as Error)?.message || "Не удалось построить план."}
            </Alert>
          )}
          {buildMutation.isSuccess && (
            <Alert tone="success">
              План успешно создан. Переключаюсь на календарный план…
            </Alert>
          )}

          <div style={{ display: "flex", justifyContent: "space-between" }}>
            {isFillMode ? (
              <Button variant="ghost" onClick={handleCancelFill}>
                ← К календарю
              </Button>
            ) : (
              <Button variant="ghost" onClick={() => setStep(2)}>
                ← Назад
              </Button>
            )}
            <Button onClick={handleSubmit} disabled={!canSubmit}>
              {buildMutation.isPending
                ? "Строим план…"
                : isFillMode
                  ? "Запустить дозаполнение"
                  : "Запустить планирование"}
            </Button>
          </div>
        </div>
      )}
    </Card>
  );
};

const thStyle = {
  textAlign: "left" as const,
  padding: "0.6rem 0.75rem",
  fontWeight: 600,
  color: "#23366f",
  fontSize: "0.9rem",
};

const tdStyle = {
  padding: "0.55rem 0.75rem",
  fontSize: "0.92rem",
  color: "#101828",
};

const subThStyle = {
  textAlign: "left" as const,
  padding: "0.4rem 0.6rem",
  fontWeight: 600,
  color: "#475467",
  fontSize: "0.82rem",
};

const subTdStyle = {
  padding: "0.4rem 0.6rem",
  fontSize: "0.86rem",
  color: "#1d2939",
};

type ExpandableKpRowProps = {
  kp: KpCandidateItem;
  isExpanded: boolean;
  isChecked: boolean;
  isIndeterminate: boolean;
  selectedCount: number;
  selectedIds: number[];
  onToggleKp: () => void;
  onToggleExpand: () => void;
  onTogglePlate: (plateId: number) => void;
};

const ExpandableKpRow = ({
  kp,
  isExpanded,
  isChecked,
  isIndeterminate,
  selectedCount,
  selectedIds,
  onToggleKp,
  onToggleExpand,
  onTogglePlate,
}: ExpandableKpRowProps) => {
  const checkboxRef = useRef<HTMLInputElement | null>(null);
  const totalPlates = kp.plates.length;

  useEffect(() => {
    if (checkboxRef.current) {
      checkboxRef.current.indeterminate = isIndeterminate;
    }
  }, [isIndeterminate]);

  const selectedSet = useMemo(() => new Set(selectedIds), [selectedIds]);

  return (
    <>
      <tr style={{ borderTop: "1px solid #e4e7ec" }}>
        <td style={{ ...tdStyle, width: 36, textAlign: "center" }}>
          <button
            type="button"
            onClick={onToggleExpand}
            aria-label={isExpanded ? "Свернуть плиты" : "Развернуть плиты"}
            aria-expanded={isExpanded}
            style={{
              border: "none",
              background: "transparent",
              cursor: "pointer",
              fontSize: "0.95rem",
              color: "#475467",
              padding: "0.1rem 0.25rem",
              lineHeight: 1,
            }}
          >
            {isExpanded ? "▾" : "▸"}
          </button>
        </td>
        <td style={tdStyle}>
          <input
            ref={checkboxRef}
            type="checkbox"
            checked={isChecked}
            onChange={onToggleKp}
            disabled={totalPlates === 0}
          />
        </td>
        <td style={tdStyle}>{kp.kp_id}</td>
        <td style={tdStyle}>{kp.customer_name || "—"}</td>
        <td style={tdStyle}>{kp.execution_terms || "—"}</td>
        <td style={tdStyle}>{kp.completion_pct.toFixed(0)}%</td>
        <td style={tdStyle}>{kp.in_plan_pct.toFixed(0)}%</td>
        <td style={tdStyle}>{kp.total_length_m.toFixed(1)}</td>
        <td style={tdStyle}>
          <span style={{ color: isIndeterminate ? "#b54708" : "#101828" }}>
            {selectedCount}/{totalPlates}
          </span>
        </td>
      </tr>
      {isExpanded && (
        <tr style={{ background: "#fafbff" }}>
          <td style={{ padding: 0 }} />
          <td colSpan={8} style={{ padding: "0.5rem 0.75rem 0.85rem" }}>
            {totalPlates === 0 ? (
              <div style={{ color: "#475467" }}>
                У этой КП нет плит со статусом «в производстве».
              </div>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    <th style={subThStyle}>Выбор</th>
                    <th style={subThStyle}>Наименование</th>
                    <th style={subThStyle}>Длина, м</th>
                    <th style={subThStyle}>Ширина, м</th>
                    <th style={subThStyle}>Кол-во</th>
                    <th style={subThStyle}>Нагрузка</th>
                  </tr>
                </thead>
                <tbody>
                  {kp.plates.map((plate) => {
                    const plateChecked = selectedSet.has(plate.id);
                    return (
                      <tr key={plate.id} style={{ borderTop: "1px solid #eef2f6" }}>
                        <td style={subTdStyle}>
                          <input
                            type="checkbox"
                            checked={plateChecked}
                            onChange={() => onTogglePlate(plate.id)}
                          />
                        </td>
                        <td style={subTdStyle}>{plate.plate_name || "—"}</td>
                        <td style={subTdStyle}>{plate.length_m.toFixed(2)}</td>
                        <td style={subTdStyle}>{plate.width_m.toFixed(2)}</td>
                        <td style={subTdStyle}>{plate.qty}</td>
                        <td style={subTdStyle}>
                          {plate.load_class !== null ? plate.load_class : "—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </td>
        </tr>
      )}
    </>
  );
};
