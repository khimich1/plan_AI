import type { ChangeEvent } from "react";
import type { CommercialDraftDetails, PlateInputMode } from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Textarea } from "@/shared/ui/Field";
import { StepLayout } from "@/shared/ui/StepLayout";

type PlateInputStepProps = {
  draft: CommercialDraftDetails | null;
  sourceText: string;
  selectedImageName: string | null;
  errorMessage: string | null;
  isPending: boolean;
  onTextChange: (value: string) => void;
  onFileChange: (file: File | null) => void;
  onSubmit: (mode: PlateInputMode) => void;
  onNext: () => void;
};

export const PlateInputStep = ({
  draft,
  sourceText,
  selectedImageName,
  errorMessage,
  isPending,
  onTextChange,
  onFileChange,
  onSubmit,
  onNext,
}: PlateInputStepProps) => {
  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    onFileChange(event.target.files?.[0] ?? null);
  };

  return (
    <StepLayout
      title="Шаг 1. Ввод плит"
      description="Вставьте текст списка плит или загрузите фото/изображение таблицы. Backend выполнит OCR, нормализацию и подготовит превью."
      footer={
        draft ? (
          <div style={{ display: "flex", gap: "0.75rem", justifyContent: "flex-end" }}>
            <Button type="button" variant="secondary" onClick={() => onSubmit("append")} disabled={isPending}>
              {isPending ? "Обработка..." : "Добавить ещё плиты"}
            </Button>
            <Button type="button" onClick={onNext}>
              Далее
            </Button>
          </div>
        ) : null
      }
    >
      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      <Card title="Источник данных" subtitle="Можно использовать текст, изображение или оба способа по очереди.">
        <div style={{ display: "grid", gap: "1rem" }}>
          <FieldWrapper label="Список плит">
            <Textarea
              value={sourceText}
              onChange={(event) => onTextChange(event.target.value)}
              placeholder={"ПБ 78-12-8п 2\nПБ 66-12-8п 4"}
            />
          </FieldWrapper>

          <FieldWrapper label="Фото / изображение таблицы" hint="Поддерживаются только изображения.">
            <input type="file" accept="image/*" onChange={handleFileChange} />
          </FieldWrapper>

          {selectedImageName && <Alert tone="info">Выбран файл: {selectedImageName}</Alert>}

          <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>
            <Button type="button" onClick={() => onSubmit(draft ? "replace" : "replace")} disabled={isPending}>
              {draft ? "Заменить список" : "Распознать / обработать"}
            </Button>
            {draft && (
              <Button type="button" variant="ghost" onClick={() => onSubmit("append")} disabled={isPending}>
                Добавить ещё плиты
              </Button>
            )}
          </div>
        </div>
      </Card>

      {draft && (
        <>
          <Card title="Нормализованный результат" subtitle="Это текст, который backend использует для расчёта.">
            <pre
              style={{
                margin: 0,
                whiteSpace: "pre-wrap",
                fontFamily: "Consolas, monospace",
                background: "#f8fafc",
                padding: "1rem",
                borderRadius: 14,
              }}
            >
              {draft.metadata.normalized_text || "Пока нет нормализованного текста."}
            </pre>
          </Card>

          <Card title="Предпросмотр обработанного списка">
            <div style={{ display: "grid", gap: "0.75rem" }}>
              <div>Позиции: {draft.order_data.length}</div>
              <div>Предупреждения: {draft.metadata.warnings.length}</div>
              <div>Нераспознанные строки: {draft.metadata.unparsed_lines.length}</div>
              <div>Широкие плиты: {draft.metadata.wide_plate_lines.length}</div>
            </div>
          </Card>
        </>
      )}
    </StepLayout>
  );
};
