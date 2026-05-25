import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
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

export const DownloadFilesSection = ({
  draft,
  isPending,
  isSchemaPending,
  onGenerate,
  onGenerateSchema,
}: DownloadFilesSectionProps) => {
  const mainFiles = draft.files.filter((file) => file.kind !== "schema");
  const schemaFile = draft.files.find((file) => file.kind === "schema");
  const showSchemaSection = mainFiles.length > 0;

  return (
    <Card
      title="Файлы КП"
      subtitle="Генерация происходит на backend. После этого файлы можно скачать по отдельности."
      actions={
        <Button type="button" variant="secondary" onClick={onGenerate} disabled={isPending}>
          {isPending ? "Генерация..." : "Сформировать файлы"}
        </Button>
      }
    >
      {mainFiles.length === 0 ? (
        <p style={{ margin: 0, color: "#475467" }}>Файлы ещё не сгенерированы.</p>
      ) : (
        <div style={{ display: "grid", gap: "0.75rem" }}>
          {mainFiles.map((file) => (
            <div
              key={file.kind}
              style={{
                display: "flex",
                justifyContent: "space-between",
                gap: "1rem",
                border: "1px solid #e4e7ec",
                borderRadius: 12,
                padding: "0.8rem 0.9rem",
              }}
            >
              <div>
                <strong>{file.display_name}</strong>
                <div style={{ color: "#475467" }}>{file.filename}</div>
              </div>
              <Button type="button" variant="ghost" onClick={() => downloadFile(file.download_url)}>
                Скачать
              </Button>
            </div>
          ))}
        </div>
      )}

      {showSchemaSection && (
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            gap: "1rem",
            border: "1px solid #e4e7ec",
            borderRadius: 12,
            padding: "0.8rem 0.9rem",
            marginTop: mainFiles.length > 0 ? "0.75rem" : 0,
          }}
        >
          <div>
            <strong>Схема раскладки (PDF)</strong>
            <div style={{ color: "#475467" }}>
              {schemaFile ? schemaFile.filename : "Генерируется отдельным запросом после основных файлов"}
            </div>
          </div>
          {schemaFile ? (
            <Button type="button" variant="ghost" onClick={() => downloadFile(schemaFile.download_url)}>
              Скачать
            </Button>
          ) : (
            <Button
              type="button"
              variant="secondary"
              onClick={onGenerateSchema}
              disabled={isSchemaPending}
            >
              {isSchemaPending ? "Формируем…" : "Сформировать схему раскладки"}
            </Button>
          )}
        </div>
      )}
    </Card>
  );
};
