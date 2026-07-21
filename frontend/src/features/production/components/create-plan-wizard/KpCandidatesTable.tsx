import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import type { ProductionEstimate } from "@/features/production/lib/productionEstimate";
import type { KpCandidateItem } from "@/features/production/types/production";
import { KpCandidateRow } from "./KpCandidateRow";
import { thStyle } from "./tableStyles";

type Props = {
  loading: boolean;
  error: boolean;
  candidates: KpCandidateItem[] | undefined;
  estimateByKpId: Map<number, ProductionEstimate | null>;
  selectedPlatesByKp: Record<number, number[]>;
  selectedPlateQtyByKp: Record<number, Record<number, number>>;
  expandedKpIds: Set<number>;
  onToggleKp: (kp: KpCandidateItem) => void;
  onToggleExpand: (kpId: number) => void;
  onTogglePlate: (kp: KpCandidateItem, plateId: number) => void;
  onSetPlateQty: (kp: KpCandidateItem, plateId: number, qty: number) => void;
};

export const KpCandidatesTable = ({
  loading,
  error,
  candidates,
  estimateByKpId,
  selectedPlatesByKp,
  selectedPlateQtyByKp,
  expandedKpIds,
  onToggleKp,
  onToggleExpand,
  onTogglePlate,
  onSetPlateQty,
}: Props) => (
  <div style={{ border: "1px solid #e4e7ec", borderRadius: 14, overflow: "hidden" }}>
    {loading && (
      <div style={{ padding: "1rem", display: "flex", gap: "0.5rem" }}>
        <Spinner /> Загрузка КП…
      </div>
    )}
    {error && (
      <div style={{ padding: "1rem" }}>
        <Alert tone="error">Не удалось загрузить список КП.</Alert>
      </div>
    )}
    {candidates && (
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
            <th style={thStyle}>≈ дор.</th>
            <th style={thStyle}>≈ дн.</th>
            <th style={thStyle}>Плиты</th>
          </tr>
        </thead>
        <tbody>
          {candidates.length === 0 && (
            <tr>
              <td
                colSpan={11}
                style={{ padding: "1rem", textAlign: "center", color: "#475467" }}
              >
                Нет КП в работе с неразмещёнными плитами.
              </td>
            </tr>
          )}
          {candidates.map((kp) => {
            const totalPlates = kp.plates.length;
            const totalQty = kp.plates.reduce((s, p) => s + p.qty, 0);
            const selected = selectedPlatesByKp[kp.kp_id];
            const selectedIds = selected ?? [];
            const selectedCount = selectedIds.length;
            const qtyByPlate = selectedPlateQtyByKp[kp.kp_id] ?? {};
            const selectedQty = selectedIds.reduce((sum, id) => {
              const plate = kp.plates.find((p) => p.id === id);
              if (!plate) {
                return sum;
              }
              return sum + (qtyByPlate[id] ?? plate.qty);
            }, 0);
            const hasPartialQty = kp.plates.some((p) => {
              if (!selectedIds.includes(p.id)) {
                return false;
              }
              return (qtyByPlate[p.id] ?? p.qty) < p.qty;
            });
            const isChecked =
              selectedCount === totalPlates &&
              totalPlates > 0 &&
              !hasPartialQty;
            const isIndeterminate =
              (selectedCount > 0 && selectedCount < totalPlates) ||
              hasPartialQty;
            const isExpanded = expandedKpIds.has(kp.kp_id);
            const rowEstimate = estimateByKpId.get(kp.kp_id) ?? null;
            return (
              <KpCandidateRow
                key={kp.kp_id}
                kp={kp}
                isExpanded={isExpanded}
                isChecked={isChecked}
                isIndeterminate={isIndeterminate}
                selectedQty={selectedQty}
                totalQty={totalQty}
                selectedIds={selectedIds}
                plateQtyById={qtyByPlate}
                rowEstimate={rowEstimate}
                onToggleKp={() => onToggleKp(kp)}
                onToggleExpand={() => onToggleExpand(kp.kp_id)}
                onTogglePlate={(plateId) => onTogglePlate(kp, plateId)}
                onSetPlateQty={(plateId, qty) => onSetPlateQty(kp, plateId, qty)}
              />
            );
          })}
        </tbody>
      </table>
    )}
  </div>
);
