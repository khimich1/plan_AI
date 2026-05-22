import type { ChangeEvent, ClipboardEvent } from "react";
import type { CommercialDraftDetails, PlateInputMode } from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { AutoResizeTextarea, FieldWrapper, Textarea } from "@/shared/ui/Field";
import { StepLayout } from "@/shared/ui/StepLayout";

type PlateInputStepProps = {
  draft: CommercialDraftDetails | null;
  sourceText: string;
  normalizedText: string;
  selectedImageName: string | null;
  recognizedImageUrl: string | null;
  recognizedImageName: string | null;
  errorMessage: string | null;
  isRecognizing: boolean;
  onTextChange: (value: string) => void;
  onNormalizedTextChange: (value: string) => void;
  onFileChange: (file: File | null) => void;
  onImagePaste: (file: File) => void;
  onRecognize: (mode: PlateInputMode) => void;
  onProcess: () => void;
  onReset: () => void;
};

const buildClipboardImageName = (type: string) => {
  const extension = type.split("/")[1]?.split("+")[0] || "png";
  const timestamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\..+/, "").replace("T", "-");
  return `clipboard-image-${timestamp}.${extension}`;
};

const createClipboardImageFile = (file: File) =>
  new File([file], buildClipboardImageName(file.type), {
    type: file.type || "image/png",
    lastModified: Date.now(),
  });

export const PlateInputStep = ({
  draft,
  sourceText,
  normalizedText,
  selectedImageName,
  recognizedImageUrl,
  recognizedImageName,
  errorMessage,
  isRecognizing,
  onTextChange,
  onNormalizedTextChange,
  onFileChange,
  onImagePaste,
  onRecognize,
  onProcess,
  onReset,
}: PlateInputStepProps) => {
  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    onFileChange(event.target.files?.[0] ?? null);
  };

  const handlePaste = (event: ClipboardEvent<HTMLDivElement>) => {
    const imageItem = Array.from(event.clipboardData.items).find((item) => item.type.startsWith("image/"));
    if (!imageItem) {
      return;
    }

    const imageFile = imageItem.getAsFile();
    if (!imageFile) {
      return;
    }

    event.preventDefault();
    onImagePaste(createClipboardImageFile(imageFile));
  };

  return (
    <StepLayout
      title="Шаг 1. Ввод плит"
      description="Вставьте текст списка плит или загрузите фото/изображение таблицы."
      footer={
        draft ? (
          <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem", flexWrap: "wrap" }}>
            <Button type="button" variant="danger" onClick={onReset}>
              Начать заново
            </Button>
            <Button type="button" variant="primary" onClick={onProcess} disabled={isRecognizing}>
              Обработать
            </Button>
          </div>
        ) : undefined
      }
    >
      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      <Card title="Источник данных" subtitle="Можно использовать текст, изображение или оба способа по очереди.">
        <div style={{ display: "grid", gap: "1rem" }} onPaste={handlePaste}>
          <FieldWrapper label="Список плит">
            <Textarea
              value={sourceText}
              onChange={(event) => onTextChange(event.target.value)}
              placeholder={"ПБ 78-12-8п 2\nПБ 66-12-8п 4"}
            />
          </FieldWrapper>

          <FieldWrapper
            label="Фото / изображение таблицы"
            hint="Поддерживаются только изображения. Можно вставить изображение из буфера обмена: Ctrl+V."
          >
            <input type="file" accept="image/*" onChange={handleFileChange} />
          </FieldWrapper>

          {selectedImageName && <Alert tone="info">Выбран файл: {selectedImageName}</Alert>}

          <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>
            <Button
              type="button"
              variant={draft ? "ghost" : "primary"}
              onClick={() => onRecognize("replace")}
              disabled={isRecognizing}
            >
              {isRecognizing ? "Распознавание..." : draft ? "Распознать (заменить)" : "Распознать"}
            </Button>
            {draft && (
              <Button type="button" variant="ghost" onClick={() => onRecognize("append")} disabled={isRecognizing}>
                Распознать и добавить
              </Button>
            )}
          </div>
        </div>
      </Card>

      {draft && (
        <div
          style={{
            display: "grid",
            gap: "1rem",
            gridTemplateColumns: recognizedImageUrl ? "repeat(auto-fit, minmax(280px, 1fr))" : "1fr",
          }}
        >
          <div style={{ display: "grid", gap: "1rem", minWidth: 0 }}>
            <Card title="Нормализованный результат" subtitle="Это текст, который backend использует для расчёта.">
              <AutoResizeTextarea
                value={normalizedText}
                onChange={(event) => onNormalizedTextChange(event.target.value)}
                placeholder="Пока нет нормализованного текста."
              />
            </Card>

            <Card title="Предпросмотр обработанного списка">
              <div style={{ display: "grid", gap: "0.75rem" }}>
                <div>Позиции: {draft.order_data.length}</div>
                <div>Предупреждения: {draft.metadata.warnings.length}</div>
                <div>Нераспознанные строки: {draft.metadata.unparsed_lines.length}</div>
                <div>Широкие плиты: {draft.metadata.wide_plate_lines.length}</div>
              </div>
            </Card>
          </div>

          {recognizedImageUrl && (
            <div style={{ position: "sticky", top: "5rem", alignSelf: "start" }}>
              <Card title="Присланное изображение" subtitle={recognizedImageName ?? undefined}>
                <img
                  src={recognizedImageUrl}
                  alt="Исходное изображение для распознавания"
                  style={{
                    width: "100%",
                    maxHeight: "70vh",
                    objectFit: "contain",
                    borderRadius: 12,
                    border: "1px solid #e4e7ec",
                    background: "#f9fafb",
                  }}
                />
              </Card>
            </div>
          )}
        </div>
      )}
    </StepLayout>
  );
};
