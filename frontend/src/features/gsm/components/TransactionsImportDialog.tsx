import { useRef, useState, type CSSProperties, type DragEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Modal } from "@/shared/ui/Modal";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import {
  amountMismatch,
  formatAmount,
  formatLiters,
  hasFileReconcileMismatch,
  litersMismatch,
  summarizeImportReport,
} from "@/features/gsm/lib/importReport";
import { useImportGsmTransactionsMutation } from "@/features/gsm/hooks/useGsmQueries";
import type { FileImportReport, TransactionImportReport } from "@/features/gsm/types/gsm";

type Props = {
  open: boolean;
  onClose: () => void;
  onImported?: (report: TransactionImportReport) => void;
};

const ACCEPT = ".xls,application/vnd.ms-excel";

const thStyle: CSSProperties = { padding: "0.45rem", textAlign: "left", borderBottom: "1px solid #eaecf0" };
const tdStyle: CSSProperties = { padding: "0.45rem", borderBottom: "1px solid #f2f4f7", verticalAlign: "top" };
const tdNumStyle: CSSProperties = { ...tdStyle, textAlign: "right", fontVariantNumeric: "tabular-nums" };
const thNumStyle: CSSProperties = { ...thStyle, textAlign: "right" };

const mismatchCell: CSSProperties = {
  ...tdStyle,
  background: "#fef3f2",
  color: "#b42318",
  fontWeight: 600,
};

const isXlsFile = (file: File): boolean => {
  const name = file.name.toLowerCase();
  return name.endsWith(".xls") || file.type.includes("excel") || file.type.includes("spreadsheet");
};

const FileReportRow = ({ file }: { file: FileImportReport }) => {
  const mismatch = hasFileReconcileMismatch(file);
  const litBad = litersMismatch(file);
  const amtBad = amountMismatch(file);
  return (
    <tr
      data-mismatch={mismatch ? "true" : "false"}
      style={mismatch ? { outline: "1px solid #fecdca", background: "#fffbfa" } : undefined}
    >
      <td style={tdStyle}>
        <div style={{ fontWeight: 600 }}>{file.filename}</div>
        {mismatch && (
          <div style={{ color: "#b42318", fontSize: "0.85rem", marginTop: 4 }}>Расхождение итогов</div>
        )}
        {file.unmatched_cards.length > 0 && (
          <div style={{ color: "#b54708", fontSize: "0.85rem", marginTop: 4 }}>
            Неизвестные карты: {file.unmatched_cards.join(", ")}
          </div>
        )}
        {file.warnings.length > 0 && (
          <ul style={{ margin: "0.35rem 0 0", paddingLeft: "1.1rem", color: "#b54708", fontSize: "0.85rem" }}>
            {file.warnings.map((w) => (
              <li key={w}>{w}</li>
            ))}
          </ul>
        )}
      </td>
      <td style={tdNumStyle}>{file.rows_total}</td>
      <td style={tdNumStyle}>{file.rows_inserted}</td>
      <td style={tdNumStyle}>{file.rows_duplicate}</td>
      <td style={litBad ? mismatchCell : tdStyle}>
        {formatLiters(file.sum_liters)}
        {file.footer_liters != null && (
          <>
            <br />
            <span style={{ fontSize: "0.85rem", opacity: 0.9 }}>итог {formatLiters(file.footer_liters)}</span>
          </>
        )}
      </td>
      <td style={amtBad ? mismatchCell : tdStyle}>
        {formatAmount(file.sum_amount)}
        {file.footer_amount != null && (
          <>
            <br />
            <span style={{ fontSize: "0.85rem", opacity: 0.9 }}>итог {formatAmount(file.footer_amount)}</span>
          </>
        )}
      </td>
    </tr>
  );
};

export const TransactionsImportDialog = ({ open, onClose, onImported }: Props) => {
  const inputRef = useRef<HTMLInputElement>(null);
  const importMutation = useImportGsmTransactionsMutation();
  const [files, setFiles] = useState<File[]>([]);
  const [dragOver, setDragOver] = useState(false);
  const [localError, setLocalError] = useState<string | null>(null);
  const [report, setReport] = useState<TransactionImportReport | null>(null);

  const resetUi = () => {
    setFiles([]);
    setLocalError(null);
    setDragOver(false);
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

  const addFiles = (list: FileList | File[]) => {
    const next = Array.from(list).filter(isXlsFile);
    if (next.length === 0) {
      setLocalError("Нужны файлы Excel (.xls) — выгрузка по картам.");
      return;
    }
    setLocalError(null);
    setReport(null);
    setFiles((prev) => {
      const byName = new Map(prev.map((f) => [f.name, f]));
      for (const f of next) {
        byName.set(f.name, f);
      }
      return Array.from(byName.values());
    });
  };

  const runImport = async () => {
    if (files.length === 0) {
      setLocalError("Выберите хотя бы один .xls файл.");
      return;
    }
    setLocalError(null);
    try {
      const result = await importMutation.mutateAsync(files);
      setReport(result);
      onImported?.(result);
    } catch (err) {
      setLocalError(formatGsmError(err));
    }
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    if (event.dataTransfer.files?.length) {
      addFiles(event.dataTransfer.files);
    }
  };

  const summary = report ? summarizeImportReport(report) : null;

  return (
    <Modal open={open} onClose={handleClose} title="Импорт транзакций" maxWidth={720}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <p style={{ margin: 0, color: "#475467", fontSize: "0.95rem" }}>
          Загрузите один или несколько .xls (файл = выгрузка по одной карте). После импорта покажем
          сверку сумм с строкой «Итоги:» по каждому файлу. Повторная загрузка безопасна: уже
          существующие операции не задвоятся.
        </p>

        <div
          role="button"
          tabIndex={0}
          aria-label="Выбрать или перетащить файлы XLS"
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
            padding: "1.5rem 1.25rem",
            textAlign: "center",
            background: dragOver ? "#eef2ff" : "#f8faff",
            cursor: importMutation.isPending ? "wait" : "pointer",
            color: "#344054",
          }}
        >
          <div style={{ fontWeight: 600, marginBottom: "0.35rem" }}>
            {importMutation.isPending ? "Импортирую…" : "Перетащите .xls сюда"}
          </div>
          <div style={{ fontSize: "0.9rem", color: "#667085" }}>
            или нажмите, чтобы выбрать несколько файлов
            {files.length > 0 && (
              <>
                <br />
                Выбрано: {files.map((f) => f.name).join(", ")}
              </>
            )}
          </div>
          <input
            ref={inputRef}
            type="file"
            accept={ACCEPT}
            multiple
            hidden
            disabled={importMutation.isPending}
            onChange={(event) => {
              if (event.target.files?.length) {
                addFiles(event.target.files);
              }
            }}
          />
        </div>

        {(localError || importMutation.isError) && (
          <Alert tone="error">{localError ?? formatGsmError(importMutation.error)}</Alert>
        )}

        {report && summary && (
          <div style={{ display: "grid", gap: "0.65rem" }}>
            <Alert tone={summary.tone}>{summary.text}</Alert>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.9rem" }}>
                <thead>
                  <tr>
                    <th style={thStyle}>Файл</th>
                    <th style={thNumStyle}>Прочитано</th>
                    <th style={thNumStyle}>Добавлено</th>
                    <th style={thNumStyle}>Уже были</th>
                    <th style={thStyle}>Литры / итог</th>
                    <th style={thStyle}>Сумма / итог</th>
                  </tr>
                </thead>
                <tbody>
                  {report.files.map((file) => (
                    <FileReportRow key={file.filename} file={file} />
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", flexWrap: "wrap" }}>
          <Button type="button" variant="ghost" onClick={handleClose} disabled={importMutation.isPending}>
            {report ? "Закрыть" : "Отмена"}
          </Button>
          {!report && (
            <Button
              type="button"
              onClick={() => void runImport()}
              disabled={importMutation.isPending || files.length === 0}
            >
              {importMutation.isPending ? "Импорт…" : `Импортировать (${files.length})`}
            </Button>
          )}
          {report && (
            <Button
              type="button"
              variant="secondary"
              onClick={() => {
                setReport(null);
                setFiles([]);
                importMutation.reset();
                if (inputRef.current) {
                  inputRef.current.value = "";
                }
              }}
            >
              Загрузить ещё
            </Button>
          )}
        </div>
      </div>
    </Modal>
  );
};
