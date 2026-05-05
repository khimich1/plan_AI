import { useEffect } from "react";
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

export const SaveOfferSection = ({
  draft,
  lastSaveResult,
  defaultExecutionTerms,
  isPending,
  onSave,
}: SaveOfferSectionProps) => {
  const form = useForm<SaveOfferFormValues>({
    resolver: zodResolver(saveOfferSchema),
    defaultValues: {
      mode: "database",
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

  return (
    <Card title="Сохранение результата" subtitle="Выберите, как сохранить подготовленное коммерческое предложение.">
      <form
        onSubmit={form.handleSubmit(async (payload) => {
          if (hasSavedOffer) {
            return;
          }
          await onSave(payload);
        })}
        style={{ display: "grid", gap: "1rem" }}
      >
        <div style={{ display: "grid", gap: "0.75rem" }}>
          {[
            { value: "database", label: "Сохранить в БД со статусом «в работе»" },
            { value: "archive", label: "Сохранить в архив" },
            { value: "skip", label: "Не сохранять" },
          ].map((item) => (
            <label
              key={item.value}
              style={{
                display: "flex",
                gap: "0.75rem",
                alignItems: "center",
                border: "1px solid #e4e7ec",
                borderRadius: 12,
                padding: "0.9rem",
              }}
            >
              <input type="radio" value={item.value} {...form.register("mode")} disabled={isSubmitDisabled} />
              <span>{item.label}</span>
            </label>
          ))}
        </div>

        {mode !== "skip" && (
          <FieldWrapper
            label="Срок изготовления"
            hint={
              mode === "archive"
                ? `${EXECUTION_TERMS_FIELD_HINT} Необязательно для архива — можно оставить пустым.`
                : EXECUTION_TERMS_FIELD_HINT
            }
            error={form.formState.errors.executionTermsInput?.message}
          >
            <Input
              {...form.register("executionTermsInput")}
              placeholder={EXECUTION_TERMS_PLACEHOLDER}
              disabled={isSubmitDisabled}
            />
          </FieldWrapper>
        )}

        <Button type="submit" variant={hasSavedOffer ? "secondary" : "primary"} disabled={isSubmitDisabled}>
          {hasSavedOffer ? "Сохранено" : isPending ? "Сохранение..." : "Подтвердить сохранение"}
        </Button>
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
