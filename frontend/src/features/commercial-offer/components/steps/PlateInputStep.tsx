import { useEffect, useState, type ChangeEvent, type ClipboardEvent, type WheelEvent } from "react";

import { filterDraftForBatchReview } from "@/features/commercial-offer/lib/batchReview";
import type {

  CommercialDraftDetails,

  OcrCorrection,

  PlateInputMode,

  WidePlateAction,

} from "@/features/commercial-offer/types/commercialOffer";

import { KpPlatePreviewPanel } from "@/features/commercial-offer/components/KpPlatePreviewPanel";

import { PlateListEditor } from "@/features/commercial-offer/components/PlateListEditor";

import { WidePlatesInlineSection } from "@/features/commercial-offer/components/WidePlatesInlineSection";

import { Alert } from "@/shared/ui/Alert";

import { Button } from "@/shared/ui/Button";

import { Card } from "@/shared/ui/Card";

import { FieldWrapper, Textarea } from "@/shared/ui/Field";

import { StepLayout } from "@/shared/ui/StepLayout";



type PlateInputStepProps = {
  draft: CommercialDraftDetails | null;
  pendingBatchReview: boolean;
  sourceText: string;
  batchReviewText: string;
  normalizedText: string;
  selectedImageName: string | null;
  recognizedImageUrl: string | null;
  recognizedImageName: string | null;
  errorMessage: string | null;
  widePlateErrorMessage?: string | null;
  isRecognizing: boolean;
  isAiProcessing?: boolean;
  isResolvingWidePlates?: boolean;
  isConfirmingBatch?: boolean;
  isProceeding?: boolean;
  widePlateDecisions?: Record<string, { action: WidePlateAction; replacementText: string }>;
  aiInstruction?: string;
  onAiInstructionChange?: (value: string) => void;
  onApplyAi?: () => void;
  onTextChange: (value: string) => void;
  onBatchReviewTextChange: (value: string) => void;
  onFileChange: (file: File | null) => void;
  onImagePaste: (file: File) => void;
  onRecognize: (mode: PlateInputMode) => void;
  onConfirmBatch: () => void;
  onFinishPlates: () => void;
  onWidePlateDecisionChange?: (lineId: string, action: WidePlateAction, replacementText: string) => void;
  onApplyWidePlates?: () => void;
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



const IMAGE_ZOOM_MIN = 0.5;
const IMAGE_ZOOM_MAX = 3;
const IMAGE_ZOOM_STEP = 0.25;

const clampImageZoom = (value: number) => Math.min(IMAGE_ZOOM_MAX, Math.max(IMAGE_ZOOM_MIN, value));

const formatImageZoom = (zoom: number) => `${Math.round(zoom * 100)}%`;

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

      return `${rowLabel}: «${afterMark}» кол-во ${item.before?.qty ?? "?"} → ${item.after?.qty ?? "?"}`;

    }

    if (item.action === "changed_mark") {

      return `${rowLabel}: «${beforeMark}» → «${afterMark}»`;

    }

    return `${rowLabel}: ${item.reason ?? item.action}`;

  });

};



export const PlateInputStep = ({
  draft,
  pendingBatchReview,
  sourceText,
  batchReviewText,
  normalizedText,
  selectedImageName,
  recognizedImageUrl,
  recognizedImageName,
  errorMessage,
  widePlateErrorMessage,
  isRecognizing,
  isAiProcessing = false,
  isResolvingWidePlates = false,
  isConfirmingBatch = false,
  isProceeding = false,
  widePlateDecisions = {},
  aiInstruction = "",
  onAiInstructionChange,
  onApplyAi,
  onTextChange,
  onBatchReviewTextChange,
  onFileChange,
  onImagePaste,
  onRecognize,
  onConfirmBatch,
  onFinishPlates,
  onWidePlateDecisionChange,
  onApplyWidePlates,
  onReset,
}: PlateInputStepProps) => {
  const [showAdditionalActions, setShowAdditionalActions] = useState(false);
  const [showSourceInput, setShowSourceInput] = useState(false);
  const [imageZoom, setImageZoom] = useState(1);
  const hasDraft = Boolean(draft);
  const isBatchReviewMode = hasDraft && pendingBatchReview;
  const batchReviewDraft = draft && isBatchReviewMode ? filterDraftForBatchReview(draft, batchReviewText) : draft;

  useEffect(() => {
    setImageZoom(1);
  }, [recognizedImageUrl]);

  const hasSourceInput = Boolean(sourceText.trim() || selectedImageName);

  const primaryRecognizeLabel = selectedImageName

    ? isRecognizing

      ? "Распознавание..."

      : "Распознать фото"

    : isRecognizing

      ? "Обработка..."

      : "Обработать текст";



  const ocrCorrections = draft?.metadata.ocr_corrections ?? [];

  const ocrCorrectionLines = formatOcrCorrections(ocrCorrections);

  const hiddenCorrectionsCount = Math.max(

    ocrCorrections.filter((item) => item.action !== "verify_failed").length - ocrCorrectionLines.length,

    0,

  );



  const hasUnresolvedWidePlates =

    Boolean(draft?.metadata.wide_plate_lines?.length) && !draft?.metadata.wide_plates_resolved;

  const canConfirmBatch =
    isBatchReviewMode && !hasUnresolvedWidePlates && !isRecognizing && !isAiProcessing && !isConfirmingBatch;
  const canFinishPlates =
    hasDraft &&
    !pendingBatchReview &&
    !hasUnresolvedWidePlates &&
    !isRecognizing &&
    !isAiProcessing &&
    !isProceeding;



  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {

    onFileChange(event.target.files?.[0] ?? null);

  };



  const handleImageWheel = (event: WheelEvent<HTMLDivElement>) => {
    if (!event.ctrlKey) {
      return;
    }
    event.preventDefault();
    const direction = event.deltaY < 0 ? 1 : -1;
    setImageZoom((current) => clampImageZoom(Number((current + direction * IMAGE_ZOOM_STEP).toFixed(2))));
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



  const sourceInputCard = (

    <Card
      title={hasDraft ? "Добавить к списку" : "Источник данных"}
      subtitle={
        hasDraft
          ? "Загрузите ещё фото или вставьте текст — позиции добавятся к текущему списку."
          : "Вставьте текст списка плит или загрузите фото таблицы."
      }
    >

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



        <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap", alignItems: "center" }}>

          {!hasDraft && (

            <Button

              type="button"

              variant="primary"

              onClick={() => onRecognize("replace")}

              disabled={isRecognizing || isAiProcessing || !hasSourceInput}

            >

              {primaryRecognizeLabel}

            </Button>

          )}



          {hasDraft && (

            <>

              <Button

                type="button"

                variant="secondary"

                onClick={() => onRecognize("append")}

                disabled={isRecognizing || isAiProcessing || !hasSourceInput}

              >

                {isRecognizing ? "Добавление..." : "Добавить к списку"}

              </Button>

              <button

                type="button"

                onClick={() => setShowAdditionalActions((open) => !open)}

                disabled={isRecognizing || isAiProcessing}

                style={{

                  border: "none",

                  background: "none",

                  color: "#175cd3",

                  cursor: "pointer",

                  padding: "0.5rem 0",

                  font: "inherit",

                }}

              >

                {showAdditionalActions ? "▾ Дополнительно" : "▸ Дополнительно"}

              </button>

            </>

          )}

        </div>



        {hasDraft && showAdditionalActions && (

          <div

            style={{

              display: "grid",

              gap: "1rem",

              border: "1px solid #e4e7ec",

              borderRadius: 12,

              padding: "1rem",

              background: "#fafafa",

            }}

          >

            <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>

              <Button

                type="button"

                variant="ghost"

                onClick={() => onRecognize("replace")}

                disabled={isRecognizing || isAiProcessing || !hasSourceInput}

              >

                {isRecognizing ? "Замена..." : "Заменить список"}

              </Button>

            </div>



            {onAiInstructionChange && onApplyAi && (

              <FieldWrapper

                label="Инструкция для помощника"

                hint="Редкий сценарий: опишите, что сделать со списком плит."

              >

                <Textarea

                  value={aiInstruction}

                  onChange={(event) => onAiInstructionChange(event.target.value)}

                  placeholder="Например: убери строки с 6п"

                />

                <div style={{ marginTop: "0.75rem" }}>

                  <Button

                    type="button"

                    variant="ghost"

                    onClick={onApplyAi}

                    disabled={isRecognizing || isAiProcessing || !aiInstruction.trim()}

                  >

                    {isAiProcessing ? "Обработка..." : "Применить инструкцию"}

                  </Button>

                </div>

              </FieldWrapper>

            )}

          </div>

        )}

      </div>

    </Card>

  );



  return (

    <StepLayout

      title="Шаг 1. Плиты"

      description={
        isBatchReviewMode
          ? "Сверьте распознанный список текущего источника с фото и нажмите «Список верен»."
          : hasDraft
            ? "Добавьте ещё плиты или перейдите к оформлению клиента."
            : "Загрузите фото или вставьте список плит для расчёта."
      }
      footer={
        hasDraft ? (
          <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem", flexWrap: "wrap", width: "100%" }}>
            <Button type="button" variant="danger" onClick={onReset}>
              Начать заново
            </Button>
            <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>
              {isBatchReviewMode && (
                <Button
                  type="button"
                  variant="primary"
                  onClick={onConfirmBatch}
                  disabled={!canConfirmBatch}
                  title={
                    hasUnresolvedWidePlates
                      ? "Сначала примите решение по позициям шире стандартной"
                      : undefined
                  }
                >
                  {isConfirmingBatch ? "Сохранение..." : "Список верен"}
                </Button>
              )}
              <Button
                type="button"
                variant={isBatchReviewMode ? "secondary" : "primary"}
                onClick={onFinishPlates}
                disabled={!canFinishPlates}
                title={
                  pendingBatchReview
                    ? "Сначала подтвердите текущий источник — «Список верен»"
                    : hasUnresolvedWidePlates
                      ? "Сначала примите решение по позициям шире стандартной"
                      : undefined
                }
              >
                {isProceeding ? "Переход..." : "Готово, далее"}
              </Button>
            </div>
          </div>
        ) : undefined
      }
    >
      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      {!hasDraft && sourceInputCard}

      {hasDraft && draft && (
        <div style={{ display: "grid", gap: "1rem" }}>
          {isBatchReviewMode && (
            <>
              {ocrCorrectionLines.length > 0 && (
                <Alert tone="warning">
                  <div>
                    При распознавании исправлено{" "}
                    {ocrCorrections.filter((item) => item.action !== "verify_failed").length} строк(и) — проверьте
                    выделенные позиции.
                  </div>
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
                  Повторная проверка распознавания не удалась — сверьте список плит с исходным фото вручную.
                </Alert>
              )}

              <div
                style={{
                  display: "grid",
                  gap: "1rem",
                  gridTemplateColumns: recognizedImageUrl ? "minmax(0, 3fr) minmax(0, 2fr)" : "minmax(0, 1fr)",
                  alignItems: "start",
                  minWidth: 0,
                }}
              >

            {recognizedImageUrl && (

              <Card

                title="Исходное фото"

                subtitle={recognizedImageName ?? undefined}

                actions={
                  <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap", alignItems: "center" }}>
                    <div
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: "0.25rem",
                        border: "1px solid #d0d5dd",
                        borderRadius: 8,
                        padding: "0.15rem",
                        background: "#ffffff",
                      }}
                    >
                      <Button
                        type="button"
                        variant="ghost"
                        onClick={() => setImageZoom((current) => clampImageZoom(current - IMAGE_ZOOM_STEP))}
                        disabled={imageZoom <= IMAGE_ZOOM_MIN}
                        title="Уменьшить"
                        style={{ minWidth: 32, padding: "0.25rem 0.5rem" }}
                      >
                        −
                      </Button>
                      <span style={{ minWidth: 48, textAlign: "center", fontSize: "0.85rem", color: "#475467" }}>
                        {formatImageZoom(imageZoom)}
                      </span>
                      <Button
                        type="button"
                        variant="ghost"
                        onClick={() => setImageZoom((current) => clampImageZoom(current + IMAGE_ZOOM_STEP))}
                        disabled={imageZoom >= IMAGE_ZOOM_MAX}
                        title="Увеличить"
                        style={{ minWidth: 32, padding: "0.25rem 0.5rem" }}
                      >
                        +
                      </Button>
                      <Button
                        type="button"
                        variant="ghost"
                        onClick={() => setImageZoom(1)}
                        title="По ширине окна"
                        style={{ padding: "0.25rem 0.5rem", fontSize: "0.85rem" }}
                      >
                        По ширине
                      </Button>
                    </div>
                    <a
                      href={recognizedImageUrl}
                      target="_blank"
                      rel="noreferrer"
                      style={{ fontSize: "0.85rem", color: "#175cd3", textDecoration: "none" }}
                    >
                      Открыть в новой вкладке
                    </a>
                  </div>
                }
              >
                <div
                  onWheel={handleImageWheel}
                  style={{
                    overflow: "auto",
                    width: "100%",
                    maxHeight: "85vh",
                    minHeight: 440,
                    borderRadius: 12,
                    border: "1px solid #e4e7ec",
                    background: "#f9fafb",
                  }}
                >
                  <img
                    src={recognizedImageUrl}
                    alt="Исходное изображение для распознавания"
                    style={{
                      display: "block",
                      width: `${imageZoom * 100}%`,
                      height: "auto",
                      maxWidth: "none",
                    }}
                  />
                </div>
                <div style={{ marginTop: "0.5rem", fontSize: "0.8rem", color: "#667085" }}>
                  Ctrl + колёсико мыши — масштаб
                </div>

              </Card>

            )}



            <Card
              title="Список плит для расчёта"
              subtitle="Сверьте позиции текущего источника с фото или текстом."
            >
              {batchReviewDraft && (
                <PlateListEditor
                  draft={batchReviewDraft}
                  value={batchReviewText}
                  onChange={onBatchReviewTextChange}
                  minHeight={recognizedImageUrl ? 440 : undefined}
                />
              )}
            </Card>
          </div>
            </>
          )}

          {onWidePlateDecisionChange && onApplyWidePlates && (

            <WidePlatesInlineSection

              draft={draft}

              decisions={widePlateDecisions}

              isPending={isResolvingWidePlates}

              errorMessage={widePlateErrorMessage}

              onDecisionChange={onWidePlateDecisionChange}

              onApply={onApplyWidePlates}

            />

          )}

          {!isBatchReviewMode && draft && (
            <KpPlatePreviewPanel draft={draft} normalizedText={normalizedText} />
          )}

          {!isBatchReviewMode ? (
            sourceInputCard
          ) : (
            <div>
              <button
                type="button"
                onClick={() => setShowSourceInput((open) => !open)}
                style={{
                  border: "none",
                  background: "none",
                  color: "#175cd3",
                  cursor: "pointer",
                  padding: 0,
                  font: "inherit",
                }}
              >
                {showSourceInput ? "▾ Скрыть добавление к списку" : "▸ Добавить к списку"}
              </button>
            </div>
          )}

          {isBatchReviewMode && showSourceInput && sourceInputCard}
        </div>
      )}

    </StepLayout>

  );

};


