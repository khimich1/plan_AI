import { useRef, useState, type CSSProperties, type DragEvent, type ReactNode } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Modal } from "@/shared/ui/Modal";
import { useImport1cMutation } from "@/features/nomenclature/hooks/useNomenclatureQueries";
import {
  NOMENCLATURE_PRODUCT_KINDS,
  PRODUCT_KIND_LABELS,
  productKindLabel,
  type AmbiguousMatchOut,
  type DisappearedMarkOut,
  type GuidWriteOut,
  type Import1cResponse,
  type NomenclatureProductKind,
  type Unmatched1COut,
} from "@/features/nomenclature/types/nomenclature";
import { getErrorMessage } from "@/shared/lib/apiError";

type Props = {
  open: boolean;
  onClose: () => void;
};

const ACCEPT =
  ".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const selectStyle: CSSProperties = {
  width: "100%",
  border: "1px solid #d0d5dd",
  borderRadius: 12,
  padding: "0.8rem 0.9rem",
  background: "#ffffff",
};

const thStyle: CSSProperties = {
  padding: "0.45rem",
  textAlign: "left",
  borderBottom: "1px solid #eaecf0",
  color: "#475467",
  fontWeight: 600,
  overflowWrap: "anywhere",
};

const tdStyle: CSSProperties = {
  padding: "0.45rem",
  borderBottom: "1px solid #f2f4f7",
  verticalAlign: "top",
  overflowWrap: "anywhere",
};

const guidCellStyle: CSSProperties = {
  ...tdStyle,
  fontFamily: "ui-monospace, monospace",
  fontSize: "0.8rem",
  overflowWrap: "anywhere",
};

const guidFieldLabel = (field: string): string => {
  if (field === "guid_1c_u") {
    return "GUID «у»";
  }
  if (field === "guid_1c") {
    return "GUID";
  }
  return field;
};

const isExcelFile = (file: File): boolean => {
  const name = file.name.toLowerCase();
  return name.endsWith(".xlsx") || name.endsWith(".xls");
};

const reportTone = (report: Import1cResponse): "success" | "warning" =>
  report.ambiguous_count > 0 || report.disappeared_count > 0 ? "warning" : "success";

const capHint = (shown: number, total: number): string =>
  total > shown ? ` (показаны ${shown} из ${total})` : ` (${total})`;

export const Import1cDialog = ({ open, onClose }: Props) => {
  const inputRef = useRef<HTMLInputElement>(null);
  const importMutation = useImport1cMutation();
  const [file, setFile] = useState<File | null>(null);
  const [productKind, setProductKind] = useState<"" | NomenclatureProductKind>("");
  const [dragOver, setDragOver] = useState(false);
  const [localError, setLocalError] = useState<string | null>(null);
  const [report, setReport] = useState<Import1cResponse | null>(null);

  const resetUi = () => {
    setFile(null);
    setProductKind("");
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

  const pickFile = (next: File) => {
    if (!isExcelFile(next)) {
      setLocalError("Нужен файл Excel (.xls или .xlsx).");
      return;
    }
    setLocalError(null);
    setReport(null);
    setFile(next);
  };

  const runImport = async () => {
    if (!file) {
      setLocalError("Выберите файл .xls или .xlsx.");
      return;
    }
    setLocalError(null);
    try {
      const result = await importMutation.mutateAsync({
        file,
        productKind: productKind || null,
      });
      setReport(result);
    } catch (err) {
      setLocalError(getErrorMessage(err));
    }
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    const next = event.dataTransfer.files?.[0];
    if (next) {
      pickFile(next);
    }
  };

  return (
    <Modal open={open} onClose={handleClose} title="Загрузка выгрузки 1С" maxWidth={720}>
      <div style={{ display: "grid", gap: "1rem", minWidth: 0, maxWidth: "100%" }}>
        {!report && (
          <p style={{ margin: 0, color: "#475467", fontSize: "0.95rem" }}>
            Загрузите полный «Прайс-лист» (.xls) или универсальный отчёт с УИД (.xlsx).
            Выгрузка с отбором — частичный режим: позиции вне файла не помечаем как исчезнувшие.
            Повторная загрузка того же файла безопасна.
          </p>
        )}

        {!report && (
          <>
            <label style={{ display: "grid", gap: "0.45rem" }}>
              <span style={{ fontWeight: 600 }}>Группа изделий</span>
              <select
                value={productKind}
                onChange={(event) =>
                  setProductKind(event.target.value as "" | NomenclatureProductKind)
                }
                disabled={importMutation.isPending}
                style={selectStyle}
                aria-label="Группа изделий"
              >
                <option value="">определить по имени файла</option>
                {NOMENCLATURE_PRODUCT_KINDS.map((kind) => (
                  <option key={kind} value={kind}>
                    {PRODUCT_KIND_LABELS[kind]}
                  </option>
                ))}
              </select>
              <span style={{ color: "#667085", fontSize: "0.9rem" }}>
                Для смешанной выгрузки ЛМ выберите марши или ступени — по имени файла группа не
                определяется.
              </span>
            </label>

            <div
              role="button"
              tabIndex={0}
              aria-label="Выбрать или перетащить файл XLS"
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
                {importMutation.isPending ? "Загружаю…" : "Перетащите .xls или .xlsx сюда"}
              </div>
              <div style={{ fontSize: "0.9rem", color: "#667085" }}>
                или нажмите, чтобы выбрать файл
                {file ? (
                  <>
                    <br />
                    Выбран: {file.name}
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
                  const next = event.target.files?.[0];
                  if (next) {
                    pickFile(next);
                  }
                }}
              />
            </div>
          </>
        )}

        {(localError || importMutation.isError) && (
          <Alert tone="error">{localError ?? getErrorMessage(importMutation.error)}</Alert>
        )}

        {report && <ImportReportSummary report={report} />}

        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", flexWrap: "wrap" }}>
          <Button type="button" variant="ghost" onClick={handleClose} disabled={importMutation.isPending}>
            {report ? "Готово" : "Отмена"}
          </Button>
          {!report && (
            <Button
              type="button"
              onClick={() => void runImport()}
              disabled={importMutation.isPending || !file}
            >
              {importMutation.isPending ? "Загрузка…" : "Загрузить"}
            </Button>
          )}
          {report && (
            <Button type="button" variant="secondary" onClick={resetUi}>
              Загрузить ещё
            </Button>
          )}
        </div>

        {report && <ImportReportLists report={report} />}
      </div>
    </Modal>
  );
};

const countRowStyle: CSSProperties = {
  display: "flex",
  justifyContent: "space-between",
  alignItems: "baseline",
  gap: "1rem",
  padding: "0.3rem 0",
  borderBottom: "1px dashed #eef2ff",
  fontSize: "0.9rem",
};

const countValueStyle: CSSProperties = {
  flexShrink: 0,
  textAlign: "right",
};

const CountRow = ({ label, value }: { label: string; value: string | number }) => (
  <div style={countRowStyle}>
    <span style={{ color: "#475467" }}>{label}</span>
    <strong style={countValueStyle}>{value}</strong>
  </div>
);

const ImportReportSummary = ({ report }: { report: Import1cResponse }) => (
  <div style={{ display: "grid", gap: "0.85rem", minWidth: 0 }}>
    <Alert tone={reportTone(report)}>{report.summary}</Alert>
    <div>
      <CountRow label="Группа" value={productKindLabel(report.product_kind)} />
      <CountRow
        label="Режим"
        value={report.mode === "partial" ? "частичная выгрузка" : "полная выгрузка"}
      />
      <CountRow label="Новые GUID" value={report.new_guids_count} />
      <CountRow label="Обновлённые GUID" value={report.updated_guids_count} />
      <CountRow label="Ждут цены" value={report.waiting_price} />
      <CountRow label="Неоднозначные" value={report.ambiguous_count} />
      <CountRow label="Исчезли из 1С" value={report.disappeared_count} />
      <CountRow label="Нет в прайсе программы" value={report.unmatched_1c_count} />
      <div style={{ ...countRowStyle, borderBottom: "none" }}>
        <span style={{ color: "#475467" }}>Без изменений</span>
        <strong style={countValueStyle}>{report.unchanged}</strong>
      </div>
    </div>
  </div>
);

const ImportReportLists = ({ report }: { report: Import1cResponse }) => (
  <div style={{ display: "grid", gap: "0.85rem", minWidth: 0 }}>
    {report.new_guids.length > 0 && (
      <GuidList
        title={`Новые GUID${capHint(report.new_guids.length, report.new_guids_count)}`}
        rows={report.new_guids}
      />
    )}
    {report.updated_guids.length > 0 && (
      <GuidList
        title={`Обновлённые GUID${capHint(report.updated_guids.length, report.updated_guids_count)}`}
        rows={report.updated_guids}
      />
    )}
    {report.ambiguous.length > 0 && (
      <AmbiguousList
        title={`Неоднозначные${capHint(report.ambiguous.length, report.ambiguous_count)}`}
        rows={report.ambiguous}
      />
    )}
    {report.disappeared.length > 0 && (
      <DisappearedList
        title={`Исчезли из 1С${capHint(report.disappeared.length, report.disappeared_count)}`}
        rows={report.disappeared}
      />
    )}
    {report.unmatched_1c.length > 0 && (
      <UnmatchedList
        title={`Нет в прайсе программы${capHint(report.unmatched_1c.length, report.unmatched_1c_count)}`}
        rows={report.unmatched_1c}
      />
    )}
  </div>
);

const ResultTable = ({
  title,
  headers,
  children,
}: {
  title: string;
  headers: string[];
  children: ReactNode;
}) => (
  <section>
    <h3 style={{ margin: "0 0 0.45rem", fontSize: "0.95rem", color: "#23366f" }}>{title}</h3>
    <div style={{ overflowX: "auto", maxWidth: "100%" }}>
      <table
        style={{
          width: "100%",
          borderCollapse: "collapse",
          fontSize: "0.88rem",
          tableLayout: "fixed",
        }}
      >
        <thead>
          <tr>
            {headers.map((header) => (
              <th key={header} style={thStyle}>
                {header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>{children}</tbody>
      </table>
    </div>
  </section>
);

const GuidList = ({ title, rows }: { title: string; rows: GuidWriteOut[] }) => (
  <ResultTable title={title} headers={["Марка", "Поле", "GUID"]}>
    {rows.map((row) => (
      <tr key={`${row.product_kind}-${row.mark}-${row.field}-${row.guid}`}>
        <td style={tdStyle}>{row.mark}</td>
        <td style={tdStyle}>{guidFieldLabel(row.field)}</td>
        <td style={guidCellStyle}>{row.guid}</td>
      </tr>
    ))}
  </ResultTable>
);

const AmbiguousList = ({ title, rows }: { title: string; rows: AmbiguousMatchOut[] }) => (
  <ResultTable title={title} headers={["Наименование", "GUID", "Заметка"]}>
    {rows.map((row) => (
      <tr key={`${row.product_kind}-${row.name}-${row.guids.join(",")}`}>
        <td style={tdStyle}>{row.mark ? `${row.name} (${row.mark})` : row.name}</td>
        <td style={guidCellStyle}>{row.guids.join(", ")}</td>
        <td style={tdStyle}>{row.note}</td>
      </tr>
    ))}
  </ResultTable>
);

const DisappearedList = ({ title, rows }: { title: string; rows: DisappearedMarkOut[] }) => (
  <ResultTable title={title} headers={["Марка", "GUID", "Статус"]}>
    {rows.map((row) => (
      <tr key={`${row.product_kind}-${row.mark}`}>
        <td style={tdStyle}>{row.mark}</td>
        <td style={guidCellStyle}>{row.guid_1c ?? "—"}</td>
        <td style={tdStyle}>{row.match_status}</td>
      </tr>
    ))}
  </ResultTable>
);

const UnmatchedList = ({ title, rows }: { title: string; rows: Unmatched1COut[] }) => (
  <ResultTable title={title} headers={["Наименование", "GUID", "Строка"]}>
    {rows.map((row) => (
      <tr key={`${row.guid}-${row.row_index}`}>
        <td style={tdStyle}>{row.name}</td>
        <td style={guidCellStyle}>{row.guid}</td>
        <td style={tdStyle}>{row.row_index}</td>
      </tr>
    ))}
  </ResultTable>
);
