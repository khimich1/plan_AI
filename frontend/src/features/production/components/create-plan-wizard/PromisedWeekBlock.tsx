import { useMemo, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import type {
  PendingPromiseExclusion,
  PromisedBlockItem,
} from "@/features/production/types/production";
import { formatRu } from "./utils";
import { tdStyle, thStyle } from "./tableStyles";

export type PromisedWeekBlockProps = {
  items: PromisedBlockItem[];
  selectedPlatesByKp: Record<number, number[]>;
  pendingExclusion: PendingPromiseExclusion | null;
  onToggleKp: (kpId: number) => void;
  onConfirmExclusion: (reason: string) => void;
  onCancelExclusion: () => void;
};

const isChecked = (
  item: PromisedBlockItem,
  selectedPlatesByKp: Record<number, number[]>,
) => (selectedPlatesByKp[item.kp_id] ?? []).length > 0;

const ReasonForm = ({
  onConfirm,
  onCancel,
}: {
  onConfirm: (reason: string) => void;
  onCancel: () => void;
}) => {
  const [reason, setReason] = useState("");
  const trimmed = reason.trim();

  return (
    <div style={{ display: "grid", gap: "0.5rem" }}>
      <label style={{ display: "grid", gap: "0.35rem" }}>
        <span style={{ fontWeight: 600, fontSize: "0.9rem" }}>
          Почему снимаете обещанное КП?
        </span>
        <textarea
          value={reason}
          onChange={(event) => setReason(event.target.value)}
          aria-label="Причина исключения обещанного КП"
          rows={3}
          style={{
            width: "100%",
            border: "1px solid #d0d5dd",
            borderRadius: 12,
            padding: "0.7rem 0.8rem",
            resize: "vertical",
          }}
        />
      </label>
      <div style={{ display: "flex", gap: "0.5rem" }}>
        <Button
          type="button"
          onClick={() => {
            if (!trimmed) {
              return;
            }
            onConfirm(trimmed);
          }}
          disabled={!trimmed}
        >
          Подтвердить причину
        </Button>
        <Button type="button" variant="secondary" onClick={onCancel}>
          Отмена
        </Button>
      </div>
    </div>
  );
};

const PromisedRow = ({
  item,
  checked,
  pending,
  onToggleKp,
  onConfirmExclusion,
  onCancelExclusion,
}: {
  item: PromisedBlockItem;
  checked: boolean;
  pending: boolean;
  onToggleKp: (kpId: number) => void;
  onConfirmExclusion: (reason: string) => void;
  onCancelExclusion: () => void;
}) => (
  <>
    <tr style={{ borderTop: "1px solid #e4e7ec" }}>
      <td style={tdStyle}>
        <input
          type="checkbox"
          checked={checked}
          onChange={() => onToggleKp(item.kp_id)}
          aria-label={
            checked
              ? `Снять обещанное КП ${item.kp_id}`
              : `Вернуть обещанное КП ${item.kp_id}`
          }
        />
      </td>
      <td style={tdStyle}>КП №{item.kp_id}</td>
      <td style={tdStyle}>{item.customer_name || "—"}</td>
      <td style={tdStyle}>{formatRu(item.promised_date)}</td>
      <td style={tdStyle}>{item.tracks} дор.</td>
      <td style={tdStyle}>
        {item.status === "overdue" ? (
          <span aria-label="Просроченное обещание">просрочено</span>
        ) : (
          "обещано"
        )}
      </td>
    </tr>
    {pending && (
      <tr style={{ background: "#fffaf8" }}>
        <td />
        <td colSpan={5} style={{ padding: "0.6rem 0.75rem 0.9rem" }}>
          <ReasonForm onConfirm={onConfirmExclusion} onCancel={onCancelExclusion} />
        </td>
      </tr>
    )}
  </>
);

const PromisedTable = ({
  items,
  selectedPlatesByKp,
  pendingExclusion,
  onToggleKp,
  onConfirmExclusion,
  onCancelExclusion,
}: PromisedWeekBlockProps) => (
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
          <th style={thStyle}>КП</th>
          <th style={thStyle}>Клиент</th>
          <th style={thStyle}>Обещать к</th>
          <th style={thStyle}>Дорожки</th>
          <th style={thStyle}>Статус</th>
        </tr>
      </thead>
      <tbody>
        {items.map((item) => (
          <PromisedRow
            key={`${item.week_start}:${item.kp_id}`}
            item={item}
            checked={isChecked(item, selectedPlatesByKp)}
            pending={pendingExclusion?.kpId === item.kp_id}
            onToggleKp={onToggleKp}
            onConfirmExclusion={onConfirmExclusion}
            onCancelExclusion={onCancelExclusion}
          />
        ))}
      </tbody>
    </table>
  </div>
);

export const PromisedWeekBlock = ({
  items,
  selectedPlatesByKp,
  pendingExclusion,
  onToggleKp,
  onConfirmExclusion,
  onCancelExclusion,
}: PromisedWeekBlockProps) => {
  const overdueItems = useMemo(
    () => items.filter((item) => item.status === "overdue"),
    [items],
  );
  const activeItems = useMemo(
    () => items.filter((item) => item.status !== "overdue"),
    [items],
  );

  return (
    <div
      data-testid="promised-week-block"
      style={{ display: "grid", gap: "0.75rem" }}
    >
      <div style={{ fontWeight: 600, color: "#23366f", fontSize: "0.95rem" }}>
        Обещано на эту неделю
      </div>

      {items.length === 0 && (
        <div style={{ color: "#475467", fontSize: "0.9rem" }}>
          Нет обещанных КП на выбранные дни.
        </div>
      )}

      {overdueItems.length > 0 && (
        <div data-testid="promised-overdue-block">
          <Alert tone="error">
            <div style={{ display: "grid", gap: "0.5rem" }}>
              <div style={{ fontWeight: 700 }}>Обещано, но не в плане</div>
              <PromisedTable
                items={overdueItems}
                selectedPlatesByKp={selectedPlatesByKp}
                pendingExclusion={pendingExclusion}
                onToggleKp={onToggleKp}
                onConfirmExclusion={onConfirmExclusion}
                onCancelExclusion={onCancelExclusion}
              />
            </div>
          </Alert>
        </div>
      )}

      {activeItems.length > 0 && (
        <div data-testid="promised-active-block">
          <PromisedTable
            items={activeItems}
            selectedPlatesByKp={selectedPlatesByKp}
            pendingExclusion={pendingExclusion}
            onToggleKp={onToggleKp}
            onConfirmExclusion={onConfirmExclusion}
            onCancelExclusion={onCancelExclusion}
          />
        </div>
      )}
    </div>
  );
};
