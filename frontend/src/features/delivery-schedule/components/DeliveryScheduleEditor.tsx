import type { CSSProperties } from "react";
import { Button } from "@/shared/ui/Button";
import { Alert } from "@/shared/ui/Alert";
import { BatchCard } from "@/features/delivery-schedule/components/BatchCard";
import { ScheduleDocumentButtons } from "@/features/delivery-schedule/components/ScheduleDocumentButtons";
import {
  computePositionSplits,
  emptyBatchDraft,
  findQtyOverflows,
  splitProgress,
  validateBatchDates,
  type BatchDraft,
  type OfferPlateForSchedule,
} from "@/features/delivery-schedule/lib/scheduleDraft";
import type { UnmatchedRowOut } from "@/features/delivery-schedule/types/deliverySchedule";

type Props = {
  plates: OfferPlateForSchedule[];
  batches: BatchDraft[];
  onBatchesChange: (batches: BatchDraft[]) => void;
  invoiceNumber: string;
  contractNumber: string;
  onInvoiceNumberChange: (value: string) => void;
  onContractNumberChange: (value: string) => void;
  readOnly?: boolean;
  validationError?: string | null;
  /** КП для кнопок шаблона/документов и импорта. */
  kpId?: number;
  /** Есть сохранённый график — можно скачивать XLSX/PDF. */
  hasSavedSchedule?: boolean;
  unmatchedRows?: UnmatchedRowOut[];
  onDismissUnmatched?: () => void;
  onImportClick?: () => void;
};

const metaInputStyle: CSSProperties = {
  width: "100%",
  border: "1px solid #d0d5dd",
  borderRadius: 10,
  padding: "0.55rem 0.7rem",
  background: "#ffffff",
  boxSizing: "border-box",
};

export const DeliveryScheduleEditor = ({
  plates,
  batches,
  onBatchesChange,
  invoiceNumber,
  contractNumber,
  onInvoiceNumberChange,
  onContractNumberChange,
  readOnly = false,
  validationError = null,
  kpId,
  hasSavedSchedule = false,
  unmatchedRows = [],
  onDismissUnmatched,
  onImportClick,
}: Props) => {
  const splits = computePositionSplits(plates, batches);
  const { allocatedPositions, totalPositions } = splitProgress(splits);
  const overflows = findQtyOverflows(plates, batches);

  const updateBatch = (key: string, next: BatchDraft) => {
    onBatchesChange(batches.map((b) => (b.key === key ? next : b)));
  };

  const removeBatch = (key: string) => {
    onBatchesChange(batches.filter((b) => b.key !== key));
  };

  const addBatch = () => {
    onBatchesChange([...batches, emptyBatchDraft(`Партия ${batches.length + 1}`)]);
  };

  return (
    <div style={{ display: "grid", gap: "1rem" }}>
      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: "0.75rem",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        <div style={{ fontWeight: 600 }}>
          Разбито {allocatedPositions} из {totalPositions} позиций
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: "0.45rem", alignItems: "center" }}>
          {kpId != null && (
            <ScheduleDocumentButtons
              kpId={kpId}
              documentsDisabled={!hasSavedSchedule}
              compact
            />
          )}
          {!readOnly && onImportClick && (
            <Button
              type="button"
              variant="secondary"
              onClick={onImportClick}
              style={{ padding: "0.45rem 0.75rem", fontSize: "0.85rem", borderRadius: 10 }}
            >
              Импорт XLSX
            </Button>
          )}
          {!readOnly && (
            <Button
              type="button"
              variant="secondary"
              onClick={addBatch}
              style={{ padding: "0.45rem 0.75rem", fontSize: "0.85rem", borderRadius: 10 }}
            >
              + Партия
            </Button>
          )}
        </div>
      </div>

      {unmatchedRows.length > 0 && (
        <Alert tone="warning">
          <div style={{ display: "grid", gap: "0.45rem" }}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                gap: "0.75rem",
                alignItems: "flex-start",
                flexWrap: "wrap",
              }}
            >
              <strong>Несматченные строки импорта ({unmatchedRows.length})</strong>
              {onDismissUnmatched && (
                <Button
                  type="button"
                  variant="ghost"
                  onClick={onDismissUnmatched}
                  style={{ padding: "0.25rem 0.55rem", fontSize: "0.8rem", borderRadius: 8 }}
                >
                  Скрыть
                </Button>
              )}
            </div>
            <ul style={{ margin: 0, paddingLeft: "1.2rem" }}>
              {unmatchedRows.map((row) => (
                <li key={`${row.row_number}-${row.reason}`} style={{ marginBottom: "0.2rem" }}>
                  Строка {row.row_number}: {row.reason}
                </li>
              ))}
            </ul>
          </div>
        </Alert>
      )}

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))",
          gap: "0.6rem",
        }}
      >
        <label style={{ display: "grid", gap: "0.35rem" }}>
          <span style={{ fontWeight: 600, fontSize: "0.9rem" }}>№ счёта</span>
          <input
            value={invoiceNumber}
            onChange={(e) => onInvoiceNumberChange(e.target.value)}
            disabled={readOnly}
            placeholder="необязательно"
            style={metaInputStyle}
          />
        </label>
        <label style={{ display: "grid", gap: "0.35rem" }}>
          <span style={{ fontWeight: 600, fontSize: "0.9rem" }}>№ договора</span>
          <input
            value={contractNumber}
            onChange={(e) => onContractNumberChange(e.target.value)}
            disabled={readOnly}
            placeholder="необязательно"
            style={metaInputStyle}
          />
        </label>
      </div>

      {validationError && <Alert tone="error">{validationError}</Alert>}
      {overflows.length > 0 && (
        <Alert tone="warning">
          Превышение по позициям:{" "}
          {overflows
            .map((o) => `«${o.plate_name}» ${o.allocated} > ${o.ordered}`)
            .join("; ")}
        </Alert>
      )}

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "minmax(220px, 0.9fr) minmax(280px, 1.4fr)",
          gap: "1rem",
          alignItems: "start",
        }}
      >
        <aside
          style={{
            border: "1px solid #e4e7ec",
            borderRadius: 14,
            padding: "0.9rem 1rem",
            background: "#f8faff",
            display: "grid",
            gap: "0.5rem",
          }}
        >
          <h3 style={{ margin: 0, fontSize: "0.95rem" }}>Позиции КП</h3>
          {splits.length === 0 ? (
            <div style={{ color: "#667085", fontSize: "0.9rem" }}>Нет позиций для разбивки.</div>
          ) : (
            splits.map((split) => {
              const done = split.ordered > 0 && split.allocated === split.ordered;
              const over = split.allocated > split.ordered;
              return (
                <div
                  key={split.plate_id}
                  style={{
                    display: "grid",
                    gap: "0.15rem",
                    padding: "0.55rem 0.65rem",
                    borderRadius: 10,
                    background: "#ffffff",
                    border: `1px solid ${over ? "#fecdca" : done ? "#abefc6" : "#e4e7ec"}`,
                  }}
                >
                  <div style={{ fontWeight: 600, fontSize: "0.9rem" }}>
                    {split.position_number != null ? `${split.position_number}. ` : ""}
                    {split.plate_name || `№${split.plate_id}`}
                  </div>
                  <div style={{ fontSize: "0.8rem", color: "#475467" }}>
                    разбито {split.allocated} из {split.ordered}
                    {split.remaining > 0 ? ` · остаток ${split.remaining}` : ""}
                    {over ? " · превышение" : ""}
                  </div>
                </div>
              );
            })
          )}
        </aside>

        <div style={{ display: "grid", gap: "0.75rem" }}>
          <h3 style={{ margin: 0, fontSize: "0.95rem" }}>Партии</h3>
          {batches.length === 0 ? (
            <div
              style={{
                padding: "1.25rem",
                border: "1px dashed #d0d5dd",
                borderRadius: 14,
                color: "#667085",
                textAlign: "center",
              }}
            >
              Партий пока нет. Добавьте первую или сохраните пустой график.
            </div>
          ) : (
            batches.map((batch) => (
              <BatchCard
                key={batch.key}
                batch={batch}
                plates={plates}
                readOnly={readOnly}
                onChange={(next) => updateBatch(batch.key, next)}
                onRemove={() => removeBatch(batch.key)}
              />
            ))
          )}
        </div>
      </div>
    </div>
  );
};

/** Клиентская проверка перед PUT: даты + Σ qty ≤ ordered. */
export const validateScheduleEditor = (
  plates: OfferPlateForSchedule[],
  batches: BatchDraft[],
): string | null => {
  for (const batch of batches) {
    if (!batch.name.trim()) {
      return "У каждой партии должно быть название";
    }
    const dateError = validateBatchDates(batch);
    if (dateError) {
      return `${batch.name || "Партия"}: ${dateError}`;
    }
  }
  const overflows = findQtyOverflows(plates, batches);
  if (overflows.length > 0) {
    const first = overflows[0];
    return `Сумма по позиции «${first.plate_name}» (${first.allocated}) превышает qty КП (${first.ordered})`;
  }
  return null;
};
