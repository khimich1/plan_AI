import { useState, type ChangeEvent, type ClipboardEvent } from "react";
import type { CommercialDraftDetails, OcrCorrection, PlateInputMode } from "@/features/commercial-offer/types/commercialOffer";
import { KpPlatePreviewPanel } from "@/features/commercial-offer/components/KpPlatePreviewPanel";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { AutoResizeTextarea, FieldWrapper, Textarea } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
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
  isAiProcessing?: boolean;
  aiInstruction?: string;
  onAiInstructionChange?: (value: string) => void;
  onApplyAi?: () => void;
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

const formatOcrCorrections = (corrections: OcrCorrection[], maxItems = 5): string[] => {
  const actionable = corrections.filter((item) => item.action !== "verify_failed");
  return actionable.slice(0, maxItems).map((item, index) => {
    const rowLabel = item.row_index != null ? `стр. ${item.row_index}` : `#${index + 1}`;
    const beforeMark = item.before?.normalized_candidate ?? "—";
    const afterMark = item.after?.normalized_candidate ?? "—";

    if (item.action === "added") {
      const qty = item.after?.qty ?? "?";
      return `${rowLabel}: добавлено «${afterMark} ${qty}»`;
    }
    if (item.action === "removed") {
      return `${rowLabel}: удалено «${beforeMark}»`;
    }
    if (item.action === "changed_qty") {
      return `${rowLabel}: «${afterMark}» qty ${item.before?.qty ?? "?"} → ${item.after?.qty ?? "?"}`;
    }
    if (item.action === "changed_mark") {
      return `${rowLabel}: «${beforeMark}» → «${afterMark}»`;
    }
    return `${rowLabel}: ${item.reason ?? item.action}`;
  });
};

export const PlateInputStep = ({
  draft,
  sourceText,
  normalizedText,
  selectedImageName,
  recognizedImageUrl,
  recognizedImageName,
  errorMessage,
  isRecognizing,
  isAiProcessing = false,
  aiInstruction = "",
  onAiInstructionChange,
  onApplyAi,
  onTextChange,
  onNormalizedTextChange,
  onFileChange,
  onImagePaste,
  onRecognize,
  onProcess,
  onReset,
}: PlateInputStepProps) => {
  const [isImageExpanded, setIsImageExpanded] = useState(false);
  const ocrCorrections = draft?.metadata.ocr_corrections ?? [];
  const ocrCorrectionLines = formatOcrCorrections(ocrCorrections);
  const hiddenCorrectionsCount = Math.max(
    ocrCorrections.filter((item) => item.action !== "verify_failed").length - ocrCorrectionLines.length,
    0,
  );

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
            <Button type="button" variant="primary" onClick={onProcess} disabled={isRecognizing || isAiProcessing}>
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
              placeholder={"ПБ 78-12-8п 2\n71-12-8 3\nПБ 66-12-8п 4"}
            />
          </FieldWrapper>

          <FieldWrapper
            label="Фото / изображение таблицы"
            hint="Поддерживаются только изображения. Можно вставить изображение из буфера обмена: Ctrl+V."
          >
            <input type="file" accept="image/*" onChange={handleFileChange} />
          </FieldWrapper>

          {selectedImageName && <Alert tone="info">Выбран файл: {selectedImageName}</Alert>}

          {draft && onAiInstructionChange && onApplyAi && (
            <FieldWrapper
              label="Инструкция для ИИ"
              hint="Опишите, что сделать со списком плит. Можно приложить фото таблицы."
            >
              <Textarea
                value={aiInstruction}
                onChange={(event) => onAiInstructionChange(event.target.value)}
                placeholder="Например: распознай таблицу с фото и замени список / убери строки с 6п"
              />
            </FieldWrapper>
          )}

          <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>
            <Button
              type="button"
              variant={draft ? "ghost" : "primary"}
              onClick={() => onRecognize("replace")}
              disabled={isRecognizing || isAiProcessing}
            >
              {isRecognizing ? "Распознавание..." : draft ? "Распознать (заменить)" : "Распознать"}
            </Button>
            {draft && (
              <Button
                type="button"
                variant="ghost"
                onClick={() => onRecognize("append")}
                disabled={isRecognizing || isAiProcessing}
              >
                Распознать и добавить
              </Button>
            )}
            {draft && onApplyAi && (
              <Button
                type="button"
                variant="primary"
                onClick={onApplyAi}
                disabled={isRecognizing || isAiProcessing || !aiInstruction.trim()}
              >
                {isAiProcessing ? "ИИ обрабатывает..." : "ИИ"}
              </Button>
            )}
          </div>
        </div>
      </Card>

      {draft && (
        <div style={{ display: "grid", gap: "1rem", gridTemplateColumns: "minmax(0, 1fr)" }}>
          <KpPlatePreviewPanel draft={draft} normalizedText={normalizedText} />

          {ocrCorrectionLines.length > 0 && (
            <Alert tone="warning">
              <div>OCR: автоисправлено {ocrCorrections.filter((item) => item.action !== "verify_failed").length} строк(и)</div>
              <ul style={{ margin: "0.5rem 0 0", paddingLeft: "1.25rem" }}>
                {ocrCorrectionLines.map((line) => (
                  <li key={line}>{line}</li>
                ))}
              </ul>
              {hiddenCorrectionsCount > 0 && <div>… и ещё {hiddenCorrectionsCount}</div>}
            </Alert>
          )}

          {draft.metadata.ocr_verify_failed && (
            <Alert tone="warning">
              Повторная проверка OCR не удалась — сверьте список плит с исходным изображением вручную.
            </Alert>
          )}

          <div
            style={{
              display: "grid",
              gap: "1rem",
              alignItems: "start",
              gridTemplateColumns: recognizedImageUrl ? "minmax(0, 0.75fr) minmax(460px, 1.25fr)" : "1fr",
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
            </div>

            {recognizedImageUrl && (
              <div style={{ position: "sticky", top: "5rem", alignSelf: "start", minWidth: 0 }}>
                <Card
                  title="Присланное изображение"
                  subtitle={
                    recognizedImageName
                      ? `${recognizedImageName} · нажмите для увеличения`
                      : "Нажмите для увеличения"
                  }
                >
                  <button
                    type="button"
                    onClick={() => setIsImageExpanded(true)}
                    aria-label="Открыть изображение в полном размере"
                    style={{
                      display: "block",
                      width: "100%",
                      padding: 0,
                      border: "none",
                      background: "transparent",
                      cursor: "zoom-in",
                    }}
                  >
                    <img
                      src={recognizedImageUrl}
                      alt="Исходное изображение для распознавания"
                      style={{
                        width: "100%",
                        maxHeight: "min(800px, 80vh)",
                        objectFit: "contain",
                        borderRadius: 12,
                        border: "1px solid #e4e7ec",
                        background: "#f9fafb",
                      }}
                    />
                  </button>
                </Card>

                <Modal
                  open={isImageExpanded}
                  onClose={() => setIsImageExpanded(false)}
                  title={recognizedImageName ?? "Присланное изображение"}
                  maxWidth={1440}
                >
                  <img
                    src={recognizedImageUrl}
                    alt="Исходное изображение для распознавания"
                    style={{
                      width: "100%",
                      maxHeight: "calc(90vh - 5rem)",
                      objectFit: "contain",
                      borderRadius: 12,
                      border: "1px solid #e4e7ec",
                      background: "#f9fafb",
                    }}
                  />
                </Modal>
              </div>
            )}
          </div>
        </div>
      )}
    </StepLayout>
  );
};
