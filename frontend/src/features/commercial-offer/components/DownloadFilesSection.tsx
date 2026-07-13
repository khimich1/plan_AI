import type { CommercialDraftDetails, FileKind } from "@/features/commercial-offer/types/commercialOffer";
import { downloadFile } from "@/shared/lib/downloadFile";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";

type DownloadFilesSectionProps = {
  draft: CommercialDraftDetails;
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

  return (
    <Card
      title="Документы для клиента"
      subtitle="Сформируйте нужные файлы и скачайте их по отдельности."
    >
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
            const file = filesByKind.get(kind);
            if (!file) {
              return null;
            }
            return (
              <FileRow
                key={file.kind}
                title={FILE_KIND_LABELS[kind]}
                filename={file.filename}
                downloadUrl={file.download_url}
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

      {hasAnyMainFile && (
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
            <Button type="button" variant="ghost" onClick={() => downloadFile(schemaFile.download_url)}>
              Скачать
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
  downloadUrl,
}: {
  title: string;
  filename: string;
  downloadUrl: string;
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
    <Button type="button" variant="ghost" onClick={() => downloadFile(downloadUrl)}>
      Скачать
    </Button>
  </div>
);
