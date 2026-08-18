import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";
import type { SubstrateRecommendation } from "@/features/production/types/production";
import { formatRu } from "./utils";
import { tdStyle, thStyle } from "./tableStyles";

export const substrateRecommendationKey = (
  rec: Pick<SubstrateRecommendation, "kp_id" | "plate_id">,
) => `${rec.kp_id}:${rec.plate_id}`;

const bySavingMDesc = (a: SubstrateRecommendation, b: SubstrateRecommendation) =>
  b.saving_m - a.saving_m;

export type SubstrateRecommendationsBlockProps = {
  recommendations: SubstrateRecommendation[];
  selectedPlatesByKp: Record<number, number[]>;
  loading?: boolean;
  errorMessage?: string | null;
  onAnalyze: () => void;
  onToggleRecommendation: (recommendation: SubstrateRecommendation) => void;
};

export const SubstrateRecommendationsBlock = ({
  recommendations,
  selectedPlatesByKp,
  loading = false,
  errorMessage = null,
  onAnalyze,
  onToggleRecommendation,
}: SubstrateRecommendationsBlockProps) => {
  const sorted = [...recommendations].sort(bySavingMDesc);

  const isChecked = (rec: SubstrateRecommendation) =>
    (selectedPlatesByKp[rec.kp_id] ?? []).includes(rec.plate_id);

  return (
    <div style={{ display: "grid", gap: "0.5rem" }}>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          gap: "0.75rem",
          flexWrap: "wrap",
        }}
      >
        <div style={{ fontWeight: 600, color: "#23366f", fontSize: "0.95rem" }}>
          Подложки из поздних КП
        </div>
        <Button
          variant="secondary"
          onClick={onAnalyze}
          disabled={loading}
        >
          Найти подложки
        </Button>
      </div>

      <div style={{ color: "#667085", fontSize: "0.85rem" }}>
        Рекомендация — преселектор. Финальный состав может отличаться
      </div>

      {loading && (
        <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
          <Spinner /> Анализируем бэклог…
        </div>
      )}

      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      {!loading && !errorMessage && sorted.length === 0 && (
        <div style={{ color: "#475467", fontSize: "0.9rem" }}>
          Нет рекомендаций по подложкам. Нажмите «Найти подложки».
        </div>
      )}

      {sorted.length > 0 && (
        <div
          style={{
            border: "1px solid #e4e7ec",
            borderRadius: 14,
            overflow: "hidden",
          }}
        >
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead style={{ background: "#f8fafc" }}>
              <tr>
                <th style={thStyle}>Выбор</th>
                <th style={thStyle}>Наименование</th>
                <th style={thStyle}>Кол-во</th>
                <th style={thStyle}>Под</th>
                <th style={thStyle}>Нужна к</th>
                <th style={thStyle}>Хранение</th>
                <th style={thStyle}>Экономия, м</th>
              </tr>
            </thead>
            <tbody>
              {sorted.map((rec) => {
                const key = substrateRecommendationKey(rec);
                const checked = isChecked(rec);
                return (
                  <tr key={key} style={{ borderTop: "1px solid #e4e7ec" }}>
                    <td style={tdStyle}>
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={() => onToggleRecommendation(rec)}
                        aria-label={`Выбрать ${rec.plate_name}`}
                      />
                    </td>
                    <td style={tdStyle}>{rec.plate_name || "—"}</td>
                    <td style={tdStyle}>{rec.qty_recommended}</td>
                    <td style={tdStyle}>{rec.under_plate_name || "—"}</td>
                    <td style={tdStyle}>{formatRu(rec.needed_by)}</td>
                    <td style={tdStyle}>{rec.storage_days}</td>
                    <td style={tdStyle}>{rec.saving_m}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};
