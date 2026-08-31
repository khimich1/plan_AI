import { useEffect, useMemo, useState, type ChangeEvent, type ClipboardEvent } from "react";

import { AiInstructionBlock } from "@/features/commercial-offer/components/AiInstructionBlock";
import { OcrWaitBanner } from "@/features/commercial-offer/components/OcrWaitBanner";
import { PlateListEditor } from "@/features/commercial-offer/components/PlateListEditor";
import { SourceImageGallery } from "@/features/commercial-offer/components/SourceImageGallery";
import { useSourceTextLint } from "@/features/commercial-offer/hooks/useSourceTextLint";
import { lintLinesToUnparsedHighlights } from "@/features/commercial-offer/lib/plateLineHighlights";
import {
  isWaitingForFirstOcrReady,
  type PageSource,
} from "@/features/commercial-offer/lib/multiPageSource";
import type { PlateInputMode, ProductType } from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper } from "@/shared/ui/Field";

export const SOURCE_LINT_PENDING_TITLE = "Проверка списка…";
export const SOURCE_LINT_ERROR_TITLE = "Не удалось проверить список";
export const SOURCE_LINT_RED_TITLE = "Исправьте красные строки";

export type SourceSubmitGate = {
  sourceText: string;
  canSubmit: boolean;
  blockReason: string | undefined;
};

type SourceInputCardProps = {
  productType: ProductType;
  hasDraft: boolean;
  sourceText: string;
  pages: PageSource[];
  activePageId: string | null;
  softCapMessage?: string | null;
  /** Append path keeps single-file input (MVP: no multi in «Дополнительно»). */
  singleFileOnly?: boolean;
  /** When false (default), thumbnail click opens lightbox before OCR. */
  recognitionStarted?: boolean;
  isRecognizing: boolean;
  isAiProcessing?: boolean;
  listLabel: string;
  placeholder: string;
  emptySubtitle: string;
  aiHint?: string;
  aiPlaceholder?: string;
  aiInstruction?: string;
  onAiInstructionChange?: (value: string) => void;
  onApplyAi?: () => void;
  onTextChange: (value: string) => void;
  onAddFiles: (files: File[]) => void;
  onRemovePage: (id: string) => void;
  onSelectPage: (id: string) => void;
  onRecognize: (mode: PlateInputMode) => void;
  onSubmitGateChange?: (gate: SourceSubmitGate) => void;
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

export function resolveSourceSubmitDisabled(
  sourceText: string,
  hasImage: boolean,
  isBusy: boolean,
  hasSourceInput: boolean,
  gate: SourceSubmitGate,
): { disabled: boolean; title?: string } {
  if (isBusy || !hasSourceInput) {
    return { disabled: true };
  }
  const lintExpected = Boolean(sourceText.trim()) && !hasImage;
  if (!lintExpected) {
    return { disabled: false };
  }
  if (gate.sourceText !== sourceText) {
    return { disabled: true, title: SOURCE_LINT_PENDING_TITLE };
  }
  return { disabled: !gate.canSubmit, title: gate.blockReason };
}

export const SourceInputCard = ({
  productType,
  hasDraft,
  sourceText,
  pages,
  activePageId,
  softCapMessage = null,
  singleFileOnly = false,
  recognitionStarted = false,
  isRecognizing,
  isAiProcessing = false,
  listLabel,
  placeholder,
  emptySubtitle,
  aiHint,
  aiPlaceholder,
  aiInstruction = "",
  onAiInstructionChange,
  onApplyAi,
  onTextChange,
  onAddFiles,
  onRemovePage,
  onSelectPage,
  onRecognize,
  onSubmitGateChange,
}: SourceInputCardProps) => {
  const [showAdditionalActions, setShowAdditionalActions] = useState(false);
  const hasImage = pages.length > 0;
  const showOcrWaitBanner = isWaitingForFirstOcrReady(recognitionStarted, pages);
  const hasSourceInput = Boolean(sourceText.trim() || hasImage);
  const lintEnabled = Boolean(sourceText.trim()) && !hasImage;
  const lint = useSourceTextLint({ text: sourceText, productType, enabled: lintEnabled });
  const hasRedLine = lint.lines.some((line) => !line.empty && !line.ok);
  const lintBlocks = lintEnabled && (lint.isPending || hasRedLine || lint.isError);
  const canSubmitSource = hasSourceInput && !isRecognizing && !isAiProcessing && !lintBlocks;
  const sourceSubmitBlockReason = !lintBlocks
    ? undefined
    : lint.isPending
      ? SOURCE_LINT_PENDING_TITLE
      : lint.isError
        ? SOURCE_LINT_ERROR_TITLE
        : SOURCE_LINT_RED_TITLE;
  const highlights = useMemo(() => lintLinesToUnparsedHighlights(lint.lines), [lint.lines]);

  useEffect(() => {
    onSubmitGateChange?.({
      sourceText,
      canSubmit: canSubmitSource,
      blockReason: sourceSubmitBlockReason,
    });
  }, [onSubmitGateChange, sourceText, canSubmitSource, sourceSubmitBlockReason]);

  const primaryRecognizeLabel = hasImage
    ? isRecognizing
      ? "Распознавание..."
      : "Распознать фото"
    : isRecognizing
      ? "Обработка..."
      : "Обработать текст";

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.target.files ?? []);
    event.target.value = "";
    if (files.length === 0) {
      return;
    }
    if (singleFileOnly) {
      onAddFiles(files.slice(0, 1));
      return;
    }
    onAddFiles(files);
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
    onAddFiles([createClipboardImageFile(imageFile)]);
  };

  return (
    <Card
      title={hasDraft ? "Добавить к списку" : "Источник данных"}
      subtitle={
        hasDraft
          ? "Загрузите ещё фото или вставьте текст — позиции добавятся к текущему списку."
          : emptySubtitle
      }
    >
      <div style={{ display: "grid", gap: "1rem" }} onPaste={handlePaste}>
        <FieldWrapper label={listLabel}>
          <PlateListEditor
            value={sourceText}
            onChange={onTextChange}
            highlights={highlights}
            placeholder={placeholder}
            minHeight={160}
            showLineNumbers
          />
        </FieldWrapper>

        <FieldWrapper
          label="Фото / изображение таблицы"
          hint={
            singleFileOnly
              ? "Поддерживаются только изображения. Можно вставить изображение из буфера обмена: Ctrl+V."
              : "Можно выбрать несколько страниц. Вставка из буфера (Ctrl+V) добавляет в конец. Распознавание — по кнопке."
          }
        >
          <input
            type="file"
            accept="image/*"
            multiple={!singleFileOnly}
            onChange={handleFileChange}
          />
        </FieldWrapper>

        <SourceImageGallery
          pages={pages}
          activeId={activePageId}
          onSelect={onSelectPage}
          onRemove={onRemovePage}
          showErrorHint
          enableLightbox={!recognitionStarted}
        />

        {softCapMessage && <Alert tone="warning">{softCapMessage}</Alert>}

        {showOcrWaitBanner && <OcrWaitBanner />}

        <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap", alignItems: "center" }}>
          {!hasDraft && (
            <Button
              type="button"
              variant="primary"
              onClick={() => onRecognize("replace")}
              disabled={!canSubmitSource}
              title={sourceSubmitBlockReason}
            >
              {primaryRecognizeLabel}
            </Button>
          )}

          {hasDraft && (
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
                disabled={!canSubmitSource}
                title={sourceSubmitBlockReason}
              >
                {isRecognizing ? "Замена..." : "Заменить список"}
              </Button>
            </div>

            {onAiInstructionChange && onApplyAi && (
              <AiInstructionBlock
                hint={aiHint}
                placeholder={aiPlaceholder}
                instruction={aiInstruction}
                onInstructionChange={onAiInstructionChange}
                onApply={onApplyAi}
                disabled={isRecognizing}
                isProcessing={isAiProcessing}
              />
            )}
          </div>
        )}
      </div>
    </Card>
  );
};
