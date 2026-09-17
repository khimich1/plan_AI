import { useRef, useState, type DragEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import {
  usePriceDeskApplyMutation,
  usePriceDeskPreviewMutation,
  usePriceDeskStatusQuery,
} from "@/features/price-desk/hooks/usePriceDeskQueries";
import type { PriceDeskPreview } from "@/features/price-desk/types/priceDesk";
import { priceDeskKindLabel } from "@/features/price-desk/types/priceDesk";
import { getErrorMessage } from "@/shared/lib/apiError";
import { Import1cDialog } from "@/features/nomenclature/components/Import1cDialog";
import { DuplicatesSection } from "@/features/price-desk/components/DuplicatesSection";
import { TasksSection } from "@/features/price-desk/components/TasksSection";

const ACCEPT = ".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const isExcelFile = (file: File): boolean => {
  const name = file.name.toLowerCase();
  return name.endsWith(".xlsx") || name.endsWith(".xls");
};

const formatPrice = (value: number | null | undefined): string => {
  if (value == null) {
    return "—";
  }
  return value.toLocaleString("ru-RU", { maximumFractionDigits: 2 });
};

export const PricesView = () => {
  const inputRef = useRef<HTMLInputElement>(null);
  const statusQuery = usePriceDeskStatusQuery();
  const previewMutation = usePriceDeskPreviewMutation();
  const applyMutation = usePriceDeskApplyMutation();
  const [dragOver, setDragOver] = useState(false);
  const [localError, setLocalError] = useState<string | null>(null);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [preview, setPreview] = useState<PriceDeskPreview | null>(null);
  const [import1cOpen, setImport1cOpen] = useState(false);

  const busy = previewMutation.isPending || applyMutation.isPending;
  const errorText =
    localError ??
    (previewMutation.isError ? getErrorMessage(previewMutation.error) : null) ??
    (applyMutation.isError ? getErrorMessage(applyMutation.error) : null);

  const resetPreview = () => {
    setLocalError(null);
    setSelectedFile(null);
    setPreview(null);
    previewMutation.reset();
    applyMutation.reset();
    if (inputRef.current) {
      inputRef.current.value = "";
    }
  };

  const runPreview = async (file: File) => {
    if (!isExcelFile(file)) {
      setLocalError("Нужен файл Excel (.xls или .xlsx)");
      return;
    }
    setLocalError(null);
    setPreview(null);
    setSelectedFile(file);
    try {
      const result = await previewMutation.mutateAsync(file);
      setPreview(result);
    } catch {
      setPreview(null);
    }
  };

  const runApply = async () => {
    if (!selectedFile || !preview) {
      return;
    }
    setLocalError(null);
    try {
      await applyMutation.mutateAsync({
        file: selectedFile,
        fileSha256: preview.file_sha256,
      });
      resetPreview();
    } catch {
      // mutation.error shown below
    }
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    const file = event.dataTransfer.files?.[0];
    if (file) {
      void runPreview(file);
    }
  };

  return (
    <main style={{ maxWidth: 960, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
      <div style={{ display: "grid", gap: "1.25rem" }}>
        <header>
          <h1 style={{ margin: 0, fontSize: "1.75rem" }}>Прайсы и 1С</h1>
          <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
            Загрузка прайса завода, выгрузка 1С, задачи по изделиям и дубли GUID.
          </p>
        </header>

        {statusQuery.isError && (
          <Alert tone="error">{getErrorMessage(statusQuery.error)}</Alert>
        )}

        <section aria-label="Текущие прайсы">
          <h2 style={{ margin: "0 0 0.75rem", fontSize: "1.1rem" }}>Сейчас в программе</h2>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.95rem" }}>
              <thead>
                <tr style={{ textAlign: "left", color: "#475467" }}>
                  <th style={{ padding: "0.5rem 0.6rem", borderBottom: "1px solid #e4e7ec" }}>Группа</th>
                  <th style={{ padding: "0.5rem 0.6rem", borderBottom: "1px solid #e4e7ec" }}>Дата прайса</th>
                  <th style={{ padding: "0.5rem 0.6rem", borderBottom: "1px solid #e4e7ec" }}>Записано</th>
                  <th style={{ padding: "0.5rem 0.6rem", borderBottom: "1px solid #e4e7ec" }}>Строк</th>
                </tr>
              </thead>
              <tbody>
                {(statusQuery.data?.groups ?? []).map((group) => (
                  <tr key={group.product_kind}>
                    <td style={{ padding: "0.55rem 0.6rem", borderBottom: "1px solid #f2f4f7" }}>
                      {priceDeskKindLabel(group.product_kind)}
                    </td>
                    <td style={{ padding: "0.55rem 0.6rem", borderBottom: "1px solid #f2f4f7" }}>
                      {group.price_list_date ?? "—"}
                    </td>
                    <td style={{ padding: "0.55rem 0.6rem", borderBottom: "1px solid #f2f4f7" }}>
                      {group.imported_at ?? "—"}
                    </td>
                    <td style={{ padding: "0.55rem 0.6rem", borderBottom: "1px solid #f2f4f7" }}>
                      {group.row_count}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        <section aria-label="Загрузка прайса завода" style={{ display: "grid", gap: "0.85rem" }}>
          <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Загрузка прайса завода</h2>
        <div
          role="button"
          tabIndex={0}
          aria-label="Выбрать или перетащить файл Excel"
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
            cursor: busy ? "wait" : "pointer",
            color: "#344054",
          }}
        >
          <div style={{ fontWeight: 600, marginBottom: "0.35rem" }}>
            {previewMutation.isPending ? "Считаю дифф…" : "Перетащите .xls или .xlsx сюда"}
          </div>
          <div style={{ fontSize: "0.9rem", color: "#667085" }}>
            или нажмите, чтобы выбрать файл
            {selectedFile ? (
              <>
                <br />
                Выбран: {selectedFile.name}
              </>
            ) : null}
          </div>
          <input
            ref={inputRef}
            type="file"
            accept={ACCEPT}
            hidden
            disabled={busy}
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file) {
                void runPreview(file);
              }
            }}
          />
        </div>

        {errorText && <Alert tone="error">{errorText}</Alert>}

        {preview && (
          <section aria-label="Дифф прайса" style={{ display: "grid", gap: "0.75rem" }}>
            <Alert tone="info">
              {priceDeskKindLabel(preview.product_kind)}
              {preview.price_list_date ? ` · дата ${preview.price_list_date}` : ""}
              {`: ${preview.parsed_rows} строк, изменено ${preview.changed}, новых ${preview.new}, пропало ${preview.missing}, без изменений ${preview.unchanged}.`}
            </Alert>
            <DiffBucket title="Изменились" items={preview.examples.changed} />
            <DiffBucket title="Новые" items={preview.examples.new} />
            <DiffBucket
              title="Пропали из файла (останутся со старой ценой)"
              items={preview.examples.missing}
            />
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button type="button" variant="ghost" onClick={resetPreview} disabled={busy}>
                Отмена
              </Button>
              <Button type="button" onClick={() => void runApply()} disabled={busy}>
                {applyMutation.isPending ? "Записываю…" : "Записать в программу"}
              </Button>
            </div>
          </section>
        )}
        </section>

        <section aria-label="Выгрузка из 1С" style={{ display: "grid", gap: "0.75rem" }}>
          <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Выгрузка из 1С</h2>
          <p style={{ margin: 0, color: "#475467" }}>
            Полный «Прайс-лист» (.xls) или универсальный отчёт с УИД (.xlsx). Выгрузка с отбором
            работает в частичном режиме — позиции вне файла не считаем исчезнувшими.
          </p>
          <div>
            <Button type="button" onClick={() => setImport1cOpen(true)}>
              Загрузить выгрузку 1С
            </Button>
          </div>
          <Import1cDialog open={import1cOpen} onClose={() => setImport1cOpen(false)} />
        </section>

        <TasksSection />
        <DuplicatesSection />
      </div>
    </main>
  );
};

const DiffBucket = ({
  title,
  items,
}: {
  title: string;
  items: PriceDeskPreview["examples"]["changed"];
}) => {
  if (items.length === 0) {
    return null;
  }
  return (
    <div>
      <h3 style={{ margin: "0 0 0.4rem", fontSize: "0.95rem" }}>{title}</h3>
      <ul style={{ margin: 0, paddingLeft: "1.1rem", color: "#344054" }}>
        {items.map((item) => (
          <li key={`${title}-${item.key}`}>
            {item.key}: {formatPrice(item.old_price)} → {formatPrice(item.new_price)}
          </li>
        ))}
      </ul>
    </div>
  );
};
