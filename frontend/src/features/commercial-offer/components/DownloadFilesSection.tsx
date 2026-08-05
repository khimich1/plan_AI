import { useCallback, useState } from "react";

import type { CommercialDraftDetails, FileKind } from "@/features/commercial-offer/types/commercialOffer";
import { httpClient } from "@/shared/api/httpClient";
import { getErrorMessage } from "@/shared/lib/apiError";
import { saveBlobAs } from "@/shared/lib/downloadFile";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";

type DownloadFilesSectionProps = {
  draft: CommercialDraftDetails;
  isSimpleKpDraft?: boolean;
  isPending: boolean;
  isSchemaPending: boolean;
  onGenerate: () => void;
  onGenerateSchema: () => void;
};

const FILE_KIND_LABELS: Record<Exclude<FileKind, "schema">, string> = {
  pdf: "PDF коммерческого предложения",
  xlsx: "Excel со списком позиций",
  breakdown: "Детальная разбивка цен",
};

export const DownloadFilesSection = ({
  draft,
  isSimpleKpDraft = false,
  isPending,
  isSchemaPending,
  onGenerate,
  onGenerateSchema,
}: DownloadFilesSectionProps) => {
  const mainFiles = draft.files.filter((file) => file.kind !== "schema");
  const schemaFile = draft.files.find((file) => file.kind === "schema");
  const filesByKind = new Map(mainFiles.map((file) => [file.kind, file]));
  const missingMainKinds = (["pdf", "xlsx"] as const).filter((kind) => !filesByKind.has(kind));
  const hasAnyMainFile = mainFiles.length > 0;
  const [downloadError, setDownloadError] = useState<string | null>(null);
  const [downloadingKey, setDownloadingKey] = useState<string | null>(null);

  const handleDownload = useCallback(async (downloadUrl: string, fallbackFilename: string, key: string) => {
    setDownloadError(null);
    setDownloadingKey(key);
    try {
      const result = await httpClient.download(downloadUrl, fallbackFilename);
      saveBlobAs(result.blob, result.filename);
    } catch (error) {
      setDownloadError(getErrorMessage(error));
    } finally {
      setDownloadingKey(null);
    }
  }, []);

  return (
    <Card
      title="Документы для клиента"
      subtitle="Сформируйте нужные файлы и скачайте их по отдельности."
    >
      {downloadError && <Alert tone="error">{downloadError}</Alert>}

      {!hasAnyMainFile ? (
        <div style={{ display: "grid", gap: "0.75rem" }}>
          <p style={{ margin: 0, color: "#475467" }}>
            Документы ещё не сформированы. Нажмите кнопку ниже — будут подготовлены PDF и Excel.
          </p>
          <Button type="button" variant="primary" onClick={onGenerate} disabled={isPending}>
            {isPending ? "Формируем документы..." : "Сформировать PDF и Excel"}
          </Button>
        </div>
      ) : (
        <div style={{ display: "grid", gap: "0.75rem" }}>
          {(["pdf", "xlsx", "breakdown"] as const).map((kind) => {
            if (isSimpleKpDraft && kind === "breakdown") {
              return null;
            }
            const file = filesByKind.get(kind);
            if (!file) {
              return null;
            }
            return (
              <FileRow
                key={file.kind}
                title={FILE_KIND_LABELS[kind]}
                filename={file.filename}
                isDownloading={downloadingKey === file.kind}
                onDownload={() => handleDownload(file.download_url, file.filename, file.kind)}
              />
            );
          })}

          {missingMainKinds.length > 0 && (
            <Button type="button" variant="secondary" onClick={onGenerate} disabled={isPending}>
              {isPending ? "Дозагрузка..." : "Дозагрузить недостающие документы"}
            </Button>
          )}
        </div>
      )}

      {hasAnyMainFile && !isSimpleKpDraft && (
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            gap: "1rem",
            border: "1px solid #e4e7ec",
            borderRadius: 12,
            padding: "0.8rem 0.9rem",
            marginTop: "0.75rem",
            alignItems: "center",
          }}
        >
          <div>
            <strong>Схема раскладки (PDF)</strong>
            <div style={{ color: "#475467" }}>
              {schemaFile ? schemaFile.filename : "Формируется отдельно после основных документов"}
            </div>
          </div>
          {schemaFile ? (
            <Button
              type="button"
              variant="ghost"
              disabled={downloadingKey === "schema"}
              onClick={() => handleDownload(schemaFile.download_url, schemaFile.filename, "schema")}
            >
              {downloadingKey === "schema" ? "Скачиваем…" : "Скачать"}
            </Button>
          ) : (
            <Button type="button" variant="secondary" onClick={onGenerateSchema} disabled={isSchemaPending}>
              {isSchemaPending ? "Формируем…" : "Сформировать схему"}
            </Button>
          )}
        </div>
      )}
    </Card>
  );
};

const FileRow = ({
  title,
  filename,
  isDownloading,
  onDownload,
}: {
  title: string;
  filename: string;
  isDownloading: boolean;
  onDownload: () => void;
}) => (
  <div
    style={{
      display: "flex",
      justifyContent: "space-between",
      gap: "1rem",
      border: "1px solid #e4e7ec",
      borderRadius: 12,
      padding: "0.8rem 0.9rem",
      alignItems: "center",
    }}
  >
    <div>
      <strong>{title}</strong>
      <div style={{ color: "#475467" }}>{filename}</div>
    </div>
    <Button type="button" variant="ghost" disabled={isDownloading} onClick={onDownload}>
      {isDownloading ? "Скачиваем…" : "Скачать"}
    </Button>
  </div>
);
