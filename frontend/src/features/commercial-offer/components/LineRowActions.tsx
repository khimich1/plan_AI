import type { LineRowHandlers } from "@/features/commercial-offer/lib/lineRowHandlers";
import { useEffect, useState, type CSSProperties, type KeyboardEvent } from "react";
import { Button } from "@/shared/ui/Button";

type LineRowActionsProps = {
  lineId: string;
  defaultQty: number;
  defaultSourceText: string;
  onSave: (payload: { qty?: number; sourceText?: string }) => void | Promise<void>;
  onDelete: () => void | Promise<void>;
  saveError?: string | null;
  busy?: boolean;
};

const iconButtonStyle: CSSProperties = {
  background: "transparent",
  border: "1px solid transparent",
  borderRadius: 8,
  padding: 4,
  cursor: "pointer",
  color: "#344054",
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  lineHeight: 0,
};

const fieldStyle: CSSProperties = {
  border: "1px solid #d0d5dd",
  borderRadius: 8,
  padding: "0.35rem 0.5rem",
  background: "#ffffff",
  font: "inherit",
};

const PencilInSquareIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path
      d="m16.862 4.487 1.687-1.688a1.875 1.875 0 1 1 2.652 2.652L10.582 16.07a4.5 4.5 0 0 1-1.897 1.13L6 18l.8-2.685a4.5 4.5 0 0 1 1.13-1.897z"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
    <path
      d="M19.5 7.125 16.862 4.487M18 14v4.75A2.25 2.25 0 0 1 15.75 21H5.25A2.25 2.25 0 0 1 3 18.75V8.25A2.25 2.25 0 0 1 5.25 6H10"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

const TrashIcon = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
    <path
      d="m14.74 9-.346 9m-4.788 0L9.26 9m9.968-3.21c.342.052.682.107 1.022.166m-1.022-.165L18.16 19.673a2.25 2.25 0 0 1-2.244 2.077H8.084a2.25 2.25 0 0 1-2.244-2.077L4.772 5.79m14.456 0a48.108 48.108 0 0 0-3.478-.397m-12 .562c.34-.059.68-.114 1.022-.165m0 0a48.11 48.11 0 0 1 3.478-.397m7.5 0v-.916c0-1.18-.91-2.164-2.09-2.201a51.964 51.964 0 0 0-3.32 0c-1.18.037-2.09 1.022-2.09 2.201v.916m7.5 0a48.667 48.667 0 0 0-7.5 0"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
    />
  </svg>
);

export const LineRowActions = ({
  lineId,
  defaultQty,
  defaultSourceText,
  onSave,
  onDelete,
  saveError,
  busy = false,
}: LineRowActionsProps) => {
  const [editing, setEditing] = useState(false);
  const [qtyDraft, setQtyDraft] = useState(String(defaultQty));
  const [sourceDraft, setSourceDraft] = useState(defaultSourceText);

  useEffect(() => {
    if (!editing) {
      setQtyDraft(String(defaultQty));
      setSourceDraft(defaultSourceText);
    }
  }, [defaultQty, defaultSourceText, editing]);

  const closeEditor = () => {
    setEditing(false);
    setQtyDraft(String(defaultQty));
    setSourceDraft(defaultSourceText);
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLElement>) => {
    if (event.key === "Escape") {
      event.preventDefault();
      closeEditor();
    }
  };

  const handleSave = () => {
    const nextQty = Number(qtyDraft);
    const nextSource = sourceDraft.trim();
    const qtyChanged = Number.isFinite(nextQty) && nextQty !== defaultQty;
    const sourceChanged = nextSource !== defaultSourceText.trim();
    if (sourceChanged) {
      void onSave({ sourceText: nextSource, ...(qtyChanged ? { qty: nextQty } : {}) });
    } else if (qtyChanged && Number.isFinite(nextQty) && nextQty >= 1) {
      void onSave({ qty: nextQty });
    }
    setEditing(false);
  };

  if (editing) {
    return (
      <div style={{ display: "grid", gap: "0.35rem", minWidth: 220 }} onKeyDown={handleKeyDown}>
        <label style={{ display: "grid", gap: 4, fontSize: "0.82rem" }}>
          Кол-во
          <input
            aria-label={`Количество строки ${lineId}`}
            type="number"
            min={1}
            value={qtyDraft}
            disabled={busy}
            onKeyDown={handleKeyDown}
            onChange={(event) => setQtyDraft(event.target.value)}
            style={fieldStyle}
          />
        </label>
        <label style={{ display: "grid", gap: 4, fontSize: "0.82rem" }}>
          Как в списке
          <input
            aria-label={`Текст строки ${lineId}`}
            value={sourceDraft}
            disabled={busy}
            onKeyDown={handleKeyDown}
            onChange={(event) => setSourceDraft(event.target.value)}
            style={fieldStyle}
          />
        </label>
        {saveError ? <div style={{ color: "#b42318", fontSize: "0.82rem" }}>{saveError}</div> : null}
        <div style={{ display: "flex", gap: "0.35rem" }}>
          <Button type="button" variant="secondary" disabled={busy} onClick={handleSave}>
            Сохранить
          </Button>
          <Button type="button" variant="ghost" disabled={busy} onClick={closeEditor}>
            Отмена
          </Button>
        </div>
      </div>
    );
  }

  return (
    <div style={{ display: "inline-flex", gap: 4, alignItems: "center" }}>
      <button
        type="button"
        aria-label={`Изменить строку ${lineId}`}
        onClick={() => setEditing(true)}
        disabled={busy}
        style={iconButtonStyle}
      >
        <PencilInSquareIcon />
      </button>
      <button
        type="button"
        aria-label={`Удалить строку ${lineId}`}
        onClick={() => void onDelete()}
        disabled={busy}
        style={iconButtonStyle}
      >
        <TrashIcon />
      </button>
      {saveError ? <div style={{ color: "#b42318", fontSize: "0.82rem" }}>{saveError}</div> : null}
    </div>
  );
};

const actionTdStyle: CSSProperties = {
  padding: "0.55rem 0.35rem",
  borderBottom: "1px solid #f2f4f7",
  whiteSpace: "nowrap",
  verticalAlign: "top",
};

export const LineActionsHeader = ({ enabled }: { enabled: boolean }) => {
  if (!enabled) {
    return null;
  }
  return (
    <th
      style={{
        padding: "0.55rem 0.65rem",
        borderBottom: "1px solid #e4e7ec",
        width: "1%",
      }}
    />
  );
};

export const LineActionsCell = ({
  handlers,
  lineId,
  qty,
  sourceText,
}: {
  handlers?: LineRowHandlers;
  lineId?: string | null;
  qty: number;
  sourceText: string;
}) => {
  if (!handlers) {
    return null;
  }
  return (
    <td style={actionTdStyle}>
      {lineId ? (
        <LineRowActions
          lineId={lineId}
          defaultQty={qty}
          defaultSourceText={sourceText}
          saveError={handlers.rowError?.lineId === lineId ? handlers.rowError.message : null}
          onSave={(payload) => void handlers.onSaveLine(lineId, payload)}
          onDelete={() => void handlers.onDeleteLine(lineId)}
        />
      ) : null}
    </td>
  );
};

