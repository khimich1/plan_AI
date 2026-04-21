import { useMemo, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import {
  useBuildPlanMutation,
  useDayOccupancyQuery,
  useKpCandidatesQuery,
} from "@/features/production/hooks/useProductionQueries";
import type { FilterMethod } from "@/features/production/types/production";

type Props = {
  onCreated?: () => void;
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

const formatISO = (date: Date) => {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
};

export const CreatePlanWizard = ({ onCreated }: Props) => {
  const [step, setStep] = useState<1 | 2 | 3>(1);
  const [startDate, setStartDate] = useState<string>(todayISO());
  const [tracksCount, setTracksCount] = useState<number>(1);
  const [filterMethod, setFilterMethod] = useState<FilterMethod>("all");
  const [selectedKpIds, setSelectedKpIds] = useState<number[]>([]);
  const [planName, setPlanName] = useState<string>("");

  const occupancyQuery = useDayOccupancyQuery();
  const candidatesQuery = useKpCandidatesQuery(step === 3 && filterMethod === "kp");
  const buildMutation = useBuildPlanMutation();

  const occupancy = occupancyQuery.data?.occupancy ?? {};
  const maxPerDay = occupancyQuery.data?.max_per_day ?? MAX_PER_DAY;

  const selectedDay = startDate;
  const occupiedOnStart = occupancy[selectedDay] ?? 0;
  const freeOnStart = Math.max(0, maxPerDay - occupiedOnStart);

  const previewDays = useMemo(() => {
    const result: Array<{ iso: string; label: string; occupied: number }> = [];
    const start = parseISODate(startDate);
    for (let i = 0; i < 10; i += 1) {
      const d = new Date(start);
      d.setDate(d.getDate() + i);
      const iso = formatISO(d);
      result.push({
        iso,
        label: d.toLocaleDateString("ru-RU", { day: "2-digit", month: "2-digit" }),
        occupied: occupancy[iso] ?? 0,
      });
    }
    return result;
  }, [startDate, occupancy]);

  const toggleKp = (kpId: number) => {
    setSelectedKpIds((prev) =>
      prev.includes(kpId) ? prev.filter((id) => id !== kpId) : [...prev, kpId],
    );
  };

  const canProceedStep1 = Boolean(startDate);
  const canProceedStep2 = tracksCount >= 1 && tracksCount <= 50;
  const canSubmit =
    canProceedStep1 &&
    canProceedStep2 &&
    (filterMethod === "all" || selectedKpIds.length > 0) &&
    !buildMutation.isPending &&
    !buildMutation.isSuccess;

  const handleSubmit = () => {
    buildMutation.mutate(
      {
        start_date: startDate,
        tracks_count: tracksCount,
        filter_method: filterMethod,
        selected_kp_ids: filterMethod === "kp" ? selectedKpIds : undefined,
        plan_name: planName.trim() ? planName.trim() : undefined,
      },
      {
        onSuccess: () => {
          setStep(1);
          setSelectedKpIds([]);
          setPlanName("");
          onCreated?.();
        },
      },
    );
  };

  return (
    <Card
      title="Начать планирование"
      subtitle="Мастер создания нового производственного плана в три шага."
    >
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

      {step === 1 && (
        <div style={{ display: "grid", gap: "1rem" }}>
          <FieldWrapper label="Дата начала планирования" hint="План построится начиная с этого дня (рабочие дни учитываются автоматически).">
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
            <div style={{ fontWeight: 600, marginBottom: "0.5rem" }}>Предварительная загрузка дат (из других планов):</div>
            {occupancyQuery.isLoading ? (
              <Spinner />
            ) : (
              <div className="prod-calendar__preview">
                {previewDays.map((day) => {
                  const state =
                    day.occupied >= maxPerDay
                      ? "prod-calendar__day--full"
                      : day.occupied > 0
                        ? "prod-calendar__day--partial"
                        : "prod-calendar__day--empty";
                  return (
                    <div key={day.iso} className={`prod-calendar__preview-cell ${state}`}>
                      <span style={{ fontWeight: 600 }}>{day.label}</span>
                      <span>{day.occupied}/{maxPerDay}</span>
                    </div>
                  );
                })}
              </div>
            )}
            <div style={{ color: "#475467", marginTop: "0.35rem" }}>
              На выбранный день уже занято <strong>{occupiedOnStart}</strong> из {maxPerDay} дорожек (свободно {freeOnStart}).
            </div>
          </div>

          <div style={{ display: "flex", justifyContent: "flex-end" }}>
            <Button onClick={() => setStep(2)} disabled={!canProceedStep1}>
              Далее →
            </Button>
          </div>
        </div>
      )}

      {step === 2 && (
        <div style={{ display: "grid", gap: "1rem" }}>
          <FieldWrapper label="Количество дорожек в день" hint="От 1 до 50. Максимум одновременно задействованных дорожек на одной дате.">
            <Input
              type="number"
              min={1}
              max={50}
              value={tracksCount}
              onChange={(e) => setTracksCount(Math.max(1, Math.min(50, Number(e.target.value) || 1)))}
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
                      <th style={thStyle}>Выбор</th>
                      <th style={thStyle}>КП №</th>
                      <th style={thStyle}>Заказчик</th>
                      <th style={thStyle}>Срок</th>
                      <th style={thStyle}>Выполнено</th>
                      <th style={thStyle}>В плане</th>
                      <th style={thStyle}>Длина, м</th>
                    </tr>
                  </thead>
                  <tbody>
                    {candidatesQuery.data.items.length === 0 && (
                      <tr>
                        <td colSpan={7} style={{ padding: "1rem", textAlign: "center", color: "#475467" }}>
                          Нет КП в работе с неразмещёнными плитами.
                        </td>
                      </tr>
                    )}
                    {candidatesQuery.data.items.map((kp) => {
                      const checked = selectedKpIds.includes(kp.kp_id);
                      return (
                        <tr key={kp.kp_id} style={{ borderTop: "1px solid #e4e7ec" }}>
                          <td style={tdStyle}>
                            <input
                              type="checkbox"
                              checked={checked}
                              onChange={() => toggleKp(kp.kp_id)}
                            />
                          </td>
                          <td style={tdStyle}>{kp.kp_id}</td>
                          <td style={tdStyle}>{kp.customer_name || "—"}</td>
                          <td style={tdStyle}>{kp.execution_terms || "—"}</td>
                          <td style={tdStyle}>{kp.completion_pct.toFixed(0)}%</td>
                          <td style={tdStyle}>{kp.in_plan_pct.toFixed(0)}%</td>
                          <td style={tdStyle}>{kp.total_length_m.toFixed(1)}</td>
                        </tr>
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
            <Button variant="ghost" onClick={() => setStep(2)}>
              ← Назад
            </Button>
            <Button onClick={handleSubmit} disabled={!canSubmit}>
              {buildMutation.isPending ? "Строим план…" : "Запустить планирование"}
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
