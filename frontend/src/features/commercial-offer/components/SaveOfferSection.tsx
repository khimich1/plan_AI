import { useEffect } from "react";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { saveOfferSchema } from "@/features/commercial-offer/schemas/commercialOffer";
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
  onSave: (payload: { mode: SaveMode; executionTermsInput: string }) => void;
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

  useEffect(() => {
    form.reset({
      mode: form.getValues("mode"),
      executionTermsInput: defaultExecutionTerms,
    });
  }, [defaultExecutionTerms, form]);

  return (
    <Card title="Сохранение результата" subtitle="Выберите, как сохранить подготовленное коммерческое предложение.">
      <form onSubmit={form.handleSubmit(onSave)} style={{ display: "grid", gap: "1rem" }}>
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
              <input type="radio" value={item.value} {...form.register("mode")} />
              <span>{item.label}</span>
            </label>
          ))}
        </div>

        {mode === "database" && (
          <FieldWrapper
            label="Срок изготовления"
            hint="Можно указать дату (`25.04.2026` или `2026-04-25`), количество дней (`14 дней`) или недель (`3 недели`)."
            error={form.formState.errors.executionTermsInput?.message}
          >
            <Input {...form.register("executionTermsInput")} placeholder="Например, 14 дней" />
          </FieldWrapper>
        )}

        <Button type="submit" disabled={isPending}>
          {isPending ? "Сохранение..." : "Подтвердить сохранение"}
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
