import { useRef, useState, type DragEvent } from "react";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { Alert } from "@/shared/ui/Alert";
import { useImportDeliveryScheduleMutation } from "@/features/delivery-schedule/hooks/useDeliveryScheduleQueries";
import type { ImportDraftResponse } from "@/features/delivery-schedule/types/deliverySchedule";
import { getErrorMessage } from "@/shared/lib/apiError";

type Props = {
  open: boolean;
  onClose: () => void;
  kpId: number;
  onImported: (result: ImportDraftResponse) => void;
};

const ACCEPT = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const isXlsxFile = (file: File): boolean => {
  const name = file.name.toLowerCase();
  return name.endsWith(".xlsx") || file.type.includes("spreadsheet");
};

export const ImportScheduleDialog = ({ open, onClose, kpId, onImported }: Props) => {
  const inputRef = useRef<HTMLInputElement>(null);
  const importMutation = useImportDeliveryScheduleMutation();
  const [dragOver, setDragOver] = useState(false);
  const [localError, setLocalError] = useState<string | null>(null);
  const [selectedName, setSelectedName] = useState<string | null>(null);

  const resetUi = () => {
    setLocalError(null);
    setSelectedName(null);
    setDragOver(false);
    importMutation.reset();
    if (inputRef.current) {
      inputRef.current.value = "";
    }
  };

  const handleClose = () => {
    if (importMutation.isPending) {
      return;
    }
    resetUi();
    onClose();
  };

  const runImport = async (file: File) => {
    if (!isXlsxFile(file)) {
      setLocalError("Нужен файл Excel (.xlsx)");
      return;
    }
    setLocalError(null);
    setSelectedName(file.name);
    try {
      const result = await importMutation.mutateAsync({ kpId, file });
      onImported(result);
      resetUi();
      onClose();
    } catch {
      // mutation.error shown below
    }
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    const file = event.dataTransfer.files?.[0];
    if (file) {
      void runImport(file);
    }
  };

  return (
    <Modal open={open} onClose={handleClose} title="Импорт графика из XLSX" maxWidth={520}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <p style={{ margin: 0, color: "#475467", fontSize: "0.95rem" }}>
          Загрузите заполненный шаблон — партии подставятся в редактор без сохранения. Несматченные
          строки покажем сверху.
        </p>

        <div
          role="button"
          tabIndex={0}
          aria-label="Выбрать или перетащить файл XLSX"
          onClick={() => inputRef.current?.click()}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") {
              event.preventDefault();
              inputRef.current?.click();
            }
          }}
          onDragEnter={(event) => {
            event.preventDefault();
            setDragOver(true);
          }}
          onDragOver={(event) => {
            event.preventDefault();
            setDragOver(true);
          }}
          onDragLeave={(event) => {
            event.preventDefault();
            setDragOver(false);
          }}
          onDrop={onDrop}
          style={{
            border: `2px dashed ${dragOver ? "#2b5cff" : "#d0d5dd"}`,
            borderRadius: 14,
            padding: "1.75rem 1.25rem",
            textAlign: "center",
            background: dragOver ? "#eef2ff" : "#f8faff",
            cursor: importMutation.isPending ? "wait" : "pointer",
            color: "#344054",
          }}
        >
          <div style={{ fontWeight: 600, marginBottom: "0.35rem" }}>
            {importMutation.isPending ? "Импортирую…" : "Перетащите .xlsx сюда"}
          </div>
          <div style={{ fontSize: "0.9rem", color: "#667085" }}>
            или нажмите, чтобы выбрать файл
            {selectedName ? (
              <>
                <br />
                Выбран: {selectedName}
              </>
            ) : null}
          </div>
          <input
            ref={inputRef}
            type="file"
            accept={ACCEPT}
            hidden
            disabled={importMutation.isPending}
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) {
                void runImport(file);
              }
            }}
          />
        </div>

        {(localError || importMutation.isError) && (
          <Alert tone="error">{localError ?? getErrorMessage(importMutation.error)}</Alert>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button type="button" variant="ghost" onClick={handleClose} disabled={importMutation.isPending}>
            Отмена
          </Button>
        </div>
      </div>
    </Modal>
  );
};
