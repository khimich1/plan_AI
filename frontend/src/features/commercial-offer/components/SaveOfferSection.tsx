import { useEffect, useState } from "react";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { saveOfferSchema } from "@/features/commercial-offer/schemas/commercialOffer";
import { EXECUTION_TERMS_FIELD_HINT, EXECUTION_TERMS_PLACEHOLDER } from "@/shared/lib/executionTerms";
import type { CommercialDraftDetails, CommercialSaveResult, SaveMode } from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input } from "@/shared/ui/Field";

type SaveOfferSectionProps = {
  draft: CommercialDraftDetails;
  lastSaveResult: CommercialSaveResult | null;
  defaultExecutionTerms: string;
  isPending: boolean;
  onSave: (payload: { mode: SaveMode; executionTermsInput: string }) => Promise<void>;
};

type SaveOfferFormValues = {
  mode: SaveMode;
  executionTermsInput: string;
};

const ALTERNATIVE_SAVE_MODES: Array<{ value: SaveMode; label: string; hint: string }> = [
  { value: "database", label: "В работе", hint: "КП сохранится со статусом «в работе» — можно вернуться к нему позже." },
  { value: "skip", label: "Пропустить", hint: "Только скачать файлы, без записи в БД." },
];

export const SaveOfferSection = ({
  draft,
  lastSaveResult,
  defaultExecutionTerms,
  isPending,
  onSave,
}: SaveOfferSectionProps) => {
  const [showAlternativeModes, setShowAlternativeModes] = useState(false);
  const form = useForm<SaveOfferFormValues>({
    resolver: zodResolver(saveOfferSchema),
    defaultValues: {
      mode: "archive",
      executionTermsInput: defaultExecutionTerms,
    },
  });

  const mode = form.watch("mode");
  const hasSavedOffer = Boolean(lastSaveResult ?? draft.saved_offer);
  const isSubmitDisabled = isPending || form.formState.isSubmitting || hasSavedOffer;

  useEffect(() => {
    form.reset({
      mode: form.getValues("mode"),
      executionTermsInput: defaultExecutionTerms,
    });
  }, [defaultExecutionTerms, form]);

  const submitSave = form.handleSubmit(async (payload) => {
    if (hasSavedOffer) {
      return;
    }
    await onSave(payload);
  });

  const handlePrimarySave = async () => {
    form.setValue("mode", "archive");
    await submitSave();
  };

  return (
    <Card
      title="Сохранение результата"
      subtitle="По умолчанию КП сохраняется в архив — для завершённых предложений."
    >
      <form onSubmit={(event) => void submitSave(event)} style={{ display: "grid", gap: "1rem" }}>
        <FieldWrapper
          label="Срок изготовления"
          hint={EXECUTION_TERMS_FIELD_HINT}
          error={form.formState.errors.executionTermsInput?.message}
        >
          <Input
            {...form.register("executionTermsInput")}
            placeholder={EXECUTION_TERMS_PLACEHOLDER}
            disabled={isSubmitDisabled || mode === "skip"}
          />
        </FieldWrapper>

        <Button
          type="button"
          variant={hasSavedOffer ? "secondary" : "primary"}
          disabled={isSubmitDisabled}
          onClick={() => void handlePrimarySave()}
        >
          {hasSavedOffer ? "Сохранено" : isPending ? "Сохранение..." : "В архив"}
        </Button>

        <div style={{ display: "grid", gap: "0.5rem" }}>
          <button
            type="button"
            onClick={() => setShowAlternativeModes((open) => !open)}
            disabled={isSubmitDisabled}
            style={{
              border: "none",
              background: "none",
              color: "#175cd3",
              cursor: isSubmitDisabled ? "not-allowed" : "pointer",
              textAlign: "left",
              padding: 0,
              font: "inherit",
              opacity: isSubmitDisabled ? 0.55 : 1,
            }}
          >
            {showAlternativeModes ? "▾ Другой вариант сохранения" : "▸ Другой вариант сохранения"}
          </button>

          {showAlternativeModes && (
            <div style={{ display: "grid", gap: "0.75rem", paddingLeft: "0.25rem" }}>
              {ALTERNATIVE_SAVE_MODES.map((item) => (
                <label
                  key={item.value}
                  style={{
                    display: "grid",
                    gap: "0.35rem",
                    border: mode === item.value ? "1px solid #84adff" : "1px solid #e4e7ec",
                    borderRadius: 12,
                    padding: "0.9rem",
                    background: mode === item.value ? "#f5f8ff" : "#ffffff",
                  }}
                >
                  <span style={{ display: "flex", gap: "0.75rem", alignItems: "center" }}>
                    <input
                      type="radio"
                      value={item.value}
                      checked={mode === item.value}
                      onChange={() => form.setValue("mode", item.value)}
                      disabled={isSubmitDisabled}
                    />
                    <strong>{item.label}</strong>
                  </span>
                  <span style={{ color: "#475467", fontSize: "0.9rem", paddingLeft: "1.6rem" }}>{item.hint}</span>
                </label>
              ))}

              {mode !== "archive" && (
                <Button type="submit" variant="secondary" disabled={isSubmitDisabled}>
                  {isPending
                    ? "Сохранение..."
                    : mode === "database"
                      ? "Сохранить в работе"
                      : "Пропустить сохранение"}
                </Button>
              )}
            </div>
          )}
        </div>
      </form>

      {(lastSaveResult ?? draft.saved_offer) && (
        <div style={{ marginTop: "1rem" }}>
          <Alert tone="success">
            {lastSaveResult?.result_card.status ?? draft.saved_offer?.status}:{" "}
            {lastSaveResult?.result_card.offer_number ?? draft.offer_identity.offer_number}
          </Alert>
        </div>
      )}
    </Card>
  );
};
