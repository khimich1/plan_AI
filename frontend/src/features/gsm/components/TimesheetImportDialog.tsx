import { useRef, useState, type CSSProperties, type DragEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Modal } from "@/shared/ui/Modal";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import { useImportTimesheetMutation } from "@/features/gsm/hooks/useGsmQueries";
import type { TimesheetImportReport } from "@/features/gsm/types/gsm";

type Props = {
  open: boolean;
  onClose: () => void;
  onImported?: (report: TimesheetImportReport) => void;
};

const PERIOD_IN_NAME = /от\s+\d{1,2}\.\d{1,2}\.\d{4}/;

const tdStyle: CSSProperties = { padding: "0.45rem", borderBottom: "1px solid #f2f4f7" };

export const TimesheetImportDialog = ({ open, onClose, onImported }: Props) => {
  const inputRef = useRef<HTMLInputElement>(null);
  const importMutation = useImportTimesheetMutation();
  const [file, setFile] = useState<File | null>(null);
  const [month, setMonth] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [localError, setLocalError] = useState<string | null>(null);
  const [report, setReport] = useState<TimesheetImportReport | null>(null);

  const needsMonth = file != null && !PERIOD_IN_NAME.test(file.name);

  const resetUi = () => {
    setFile(null);
    setMonth("");
    setDragOver(false);
    setLocalError(null);
    setReport(null);
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

  const takeFile = (next: File | null) => {
    if (!next) {
      return;
    }
    if (!next.name.toLowerCase().endsWith(".xlsx")) {
      setLocalError("Нужен файл .xlsx формы Т-13.");
      return;
    }
    setLocalError(null);
    setReport(null);
    setFile(next);
  };

  const runImport = async () => {
    if (!file) {
      setLocalError("Выберите файл табеля .xlsx.");
      return;
    }
    setLocalError(null);
    try {
      const result = await importMutation.mutateAsync({
        file,
        month: needsMonth ? month || undefined : undefined,
      });
      setReport(result);
      onImported?.(result);
    } catch (err) {
      setLocalError(formatGsmError(err));
    }
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    takeFile(event.dataTransfer.files?.[0] ?? null);
  };

  return (
    <Modal open={open} onClose={handleClose} title="Импорт табеля Т-13" maxWidth={640}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <p style={{ margin: 0, color: "#475467", fontSize: "0.95rem" }}>
          Загрузите табель рабочего времени за месяц. Если в имени файла нет даты «от ДД.ММ.ГГГГ»,
          выберите месяц вручную. Повторная загрузка того же месяца обновляет явки.
        </p>

        <div
          role="button"
          tabIndex={0}
          aria-label="Выбрать или перетащить файл табеля"
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
            padding: "1.25rem",
            textAlign: "center",
            background: dragOver ? "#f5f8ff" : "#fcfcfd",
            cursor: "pointer",
          }}
        >
          {file ? file.name : "Перетащите .xlsx или нажмите, чтобы выбрать"}
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            hidden
            onChange={(event) => takeFile(event.target.files?.[0] ?? null)}
          />
        </div>

        {needsMonth && (
          <label style={{ display: "grid", gap: "0.35rem", color: "#344054" }}>
            Месяц табеля
            <input
              aria-label="Месяц табеля"
              type="month"
              value={month}
              onChange={(event) => setMonth(event.target.value)}
            />
          </label>
        )}

        {localError && <Alert tone="error">{localError}</Alert>}

        {report && (
          <div>
            <p style={{ margin: "0 0 0.5rem" }}>
              {report.period}: сопоставлено {report.matched} из {report.rows_total}.
            </p>
            {report.unmatched.length > 0 && (
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.9rem" }}>
                <thead>
                  <tr>
                    <th style={{ ...tdStyle, textAlign: "left" }}>Строка без водителя</th>
                    <th style={{ ...tdStyle, textAlign: "left" }}>Табельный</th>
                  </tr>
                </thead>
                <tbody>
                  {report.unmatched.map((row) => (
                    <tr key={`${row.full_name}-${row.personnel_number ?? ""}`}>
                      <td style={tdStyle}>{row.full_name}</td>
                      <td style={tdStyle}>{row.personnel_number ?? "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button type="button" variant="secondary" onClick={handleClose} disabled={importMutation.isPending}>
            Закрыть
          </Button>
          <Button type="button" onClick={() => void runImport()} disabled={importMutation.isPending || !file}>
            Загрузить
          </Button>
        </div>
      </div>
    </Modal>
  );
};
