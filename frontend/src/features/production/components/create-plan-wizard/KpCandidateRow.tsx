import { useEffect, useMemo, useRef } from "react";
import { Input } from "@/shared/ui/Field";
import type { KpCandidateItem } from "@/features/production/types/production";
import type { ProductionEstimate } from "@/features/production/lib/productionEstimate";
import { subTdStyle, subThStyle, tdStyle } from "./tableStyles";

export type KpCandidateRowProps = {
  kp: KpCandidateItem;
  isExpanded: boolean;
  isChecked: boolean;
  isIndeterminate: boolean;
  selectedQty: number;
  totalQty: number;
  selectedIds: number[];
  plateQtyById: Record<number, number>;
  rowEstimate: ProductionEstimate | null;
  onToggleKp: () => void;
  onToggleExpand: () => void;
  onTogglePlate: (plateId: number) => void;
  onSetPlateQty: (plateId: number, qty: number) => void;
};

export const KpCandidateRow = ({
  kp,
  isExpanded,
  isChecked,
  isIndeterminate,
  selectedQty,
  totalQty,
  selectedIds,
  plateQtyById,
  rowEstimate,
  onToggleKp,
  onToggleExpand,
  onTogglePlate,
  onSetPlateQty,
}: KpCandidateRowProps) => {
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
          {rowEstimate ? `~${rowEstimate.estimated_tracks}` : "—"}
        </td>
        <td style={tdStyle}>
          {rowEstimate ? `~${rowEstimate.estimated_days}` : "—"}
        </td>
        <td style={tdStyle}>
          <span style={{ color: isIndeterminate ? "#b54708" : "#101828" }}>
            {selectedQty}/{totalQty}
          </span>
        </td>
      </tr>
      {isExpanded && (
        <tr style={{ background: "#fafbff" }}>
          <td style={{ padding: 0 }} />
          <td colSpan={10} style={{ padding: "0.5rem 0.75rem 0.85rem" }}>
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
                    const displayQty = plateQtyById[plate.id] ?? plate.qty;
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
                        <td style={subTdStyle}>
                          <Input
                            type="number"
                            min={1}
                            max={plate.qty}
                            value={displayQty}
                            disabled={!plateChecked}
                            onChange={(e) =>
                              onSetPlateQty(plate.id, Number(e.target.value))
                            }
                          />
                        </td>
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
