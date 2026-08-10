import type { CSSProperties } from "react";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper } from "@/shared/ui/Field";
import { BatchStatusChip } from "@/features/delivery-schedule/components/BatchStatusChip";
import type { BatchDraft, OfferPlateForSchedule } from "@/features/delivery-schedule/lib/scheduleDraft";

type Props = {
  batch: BatchDraft;
  plates: OfferPlateForSchedule[];
  readOnly?: boolean;
  onChange: (next: BatchDraft) => void;
  onRemove: () => void;
};

const inputStyle: CSSProperties = {
  width: "100%",
  border: "1px solid #d0d5dd",
  borderRadius: 10,
  padding: "0.55rem 0.7rem",
  background: "#ffffff",
  boxSizing: "border-box",
};

export const BatchCard = ({ batch, plates, readOnly = false, onChange, onRemove }: Props) => {
  const plateById = new Map(plates.map((p) => [p.id, p]));

  const setField = <K extends keyof BatchDraft>(key: K, value: BatchDraft[K]) => {
    onChange({ ...batch, [key]: value });
  };

  const updateItemQty = (plateId: number, qtyRaw: string) => {
    const qty = Math.max(0, Math.trunc(Number(qtyRaw) || 0));
    const others = batch.items.filter((item) => item.plate_id !== plateId);
    const nextItems = qty > 0 ? [...others, { plate_id: plateId, qty }] : others;
    onChange({ ...batch, items: nextItems });
  };

  const qtyFor = (plateId: number): number =>
    batch.items.find((item) => item.plate_id === plateId)?.qty ?? 0;

  return (
    <article
      style={{
        border: "1px solid #e4e7ec",
        borderRadius: 14,
        padding: "0.9rem 1rem",
        background: "#ffffff",
        display: "grid",
        gap: "0.75rem",
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem", flexWrap: "wrap" }}>
        <div style={{ flex: "1 1 200px", minWidth: 0 }}>
          <FieldWrapper label="Название партии">
            <input
              value={batch.name}
              onChange={(e) => setField("name", e.target.value)}
              placeholder="Например, 1 этаж, 2 подъезд"
              disabled={readOnly}
              style={inputStyle}
            />
          </FieldWrapper>
        </div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: "0.4rem" }}>
          <BatchStatusChip status={batch.status} hint={batch.hint} readyDate={batch.ready_date} />
          {batch.changed && (
            <span style={{ fontSize: "0.75rem", color: "#b54708" }}>кол-во позиции КП изменилось</span>
          )}
          {!readOnly && (
            <Button type="button" variant="ghost" onClick={onRemove} style={{ padding: "0.35rem 0.6rem", fontSize: "0.8rem" }}>
              Удалить
            </Button>
          )}
        </div>
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))",
          gap: "0.6rem",
        }}
      >
        <FieldWrapper label="Поставка с">
          <input
            type="date"
            value={batch.deliver_from}
            onChange={(e) => setField("deliver_from", e.target.value)}
            disabled={readOnly}
            style={inputStyle}
          />
        </FieldWrapper>
        <FieldWrapper label="Поставка по">
          <input
            type="date"
            value={batch.deliver_to}
            onChange={(e) => setField("deliver_to", e.target.value)}
            disabled={readOnly}
            style={inputStyle}
          />
        </FieldWrapper>
        <FieldWrapper label="Произвести до">
          <input
            type="date"
            value={batch.produce_by}
            onChange={(e) => setField("produce_by", e.target.value)}
            disabled={readOnly}
            style={inputStyle}
          />
        </FieldWrapper>
      </div>

      <div>
        <div style={{ fontWeight: 600, marginBottom: "0.45rem", fontSize: "0.9rem" }}>Позиции в партии</div>
        {plates.length === 0 ? (
          <div style={{ color: "#667085", fontSize: "0.9rem" }}>Нет позиций КП с id.</div>
        ) : (
          <div style={{ display: "grid", gap: "0.4rem" }}>
            {plates.map((plate) => {
              const qty = qtyFor(plate.id);
              return (
                <div
                  key={plate.id}
                  style={{
                    display: "grid",
                    gridTemplateColumns: "1fr 90px",
                    gap: "0.5rem",
                    alignItems: "center",
                  }}
                >
                  <span style={{ fontSize: "0.9rem", color: qty > 0 ? "#101828" : "#667085" }}>
                    {plate.position_number != null ? `${plate.position_number}. ` : ""}
                    {plate.plate_name || plateById.get(plate.id)?.plate_name || `№${plate.id}`}
                  </span>
                  <input
                    type="number"
                    min={0}
                    step={1}
                    value={qty}
                    onChange={(e) => updateItemQty(plate.id, e.target.value)}
                    disabled={readOnly}
                    aria-label={`Кол-во ${plate.plate_name}`}
                    style={{ ...inputStyle, padding: "0.4rem 0.5rem" }}
                  />
                </div>
              );
            })}
          </div>
        )}
      </div>
    </article>
  );
};
