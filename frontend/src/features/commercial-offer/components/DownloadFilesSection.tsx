import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import { downloadFile } from "@/shared/lib/downloadFile";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";

type DownloadFilesSectionProps = {
  draft: CommercialDraftDetails;
  isPending: boolean;
  onGenerate: () => void;
};

export const DownloadFilesSection = ({ draft, isPending, onGenerate }: DownloadFilesSectionProps) => (
  <Card
    title="Файлы КП"
    subtitle="Генерация происходит на backend. После этого файлы можно скачать по отдельности."
    actions={
      <Button type="button" variant="secondary" onClick={onGenerate} disabled={isPending}>
        {isPending ? "Генерация..." : "Сформировать файлы"}
      </Button>
    }
  >
    {draft.files.length === 0 ? (
      <p style={{ margin: 0, color: "#475467" }}>Файлы ещё не сгенерированы.</p>
    ) : (
      <div style={{ display: "grid", gap: "0.75rem" }}>
        {draft.files.map((file) => (
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
  </Card>
);
