import { useEffect } from "react";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { clientConditionsSchema } from "@/features/commercial-offer/schemas/commercialOffer";
import type { ConditionsMode } from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input, Textarea } from "@/shared/ui/Field";
import { StepLayout } from "@/shared/ui/StepLayout";

type ClientConditionsStepProps = {
  defaultValues: {
    clientName: string;
    discountPercent: number;
    conditionsMode: ConditionsMode;
    deliveryConditions: string;
    paymentConditions: string;
  };
  errorMessage: string | null;
  isPending: boolean;
  onBack: () => void;
  onSubmit: (payload: {
    clientName: string;
    discountPercent: number;
    conditionsMode: ConditionsMode;
    deliveryConditions: string;
    paymentConditions: string;
  }) => void;
};

type ClientConditionsFormValues = {
  clientName: string;
  discountPercent: number;
  conditionsMode: ConditionsMode;
  deliveryConditions: string;
  paymentConditions: string;
};

export const ClientConditionsStep = ({
  defaultValues,
  errorMessage,
  isPending,
  onBack,
  onSubmit,
}: ClientConditionsStepProps) => {
  const form = useForm<ClientConditionsFormValues>({
    resolver: zodResolver(clientConditionsSchema),
    defaultValues,
  });

  useEffect(() => {
    form.reset(defaultValues);
  }, [defaultValues, form]);

  const conditionsMode = form.watch("conditionsMode");

  return (
    <StepLayout
      title="Шаг 4. Клиент и условия"
      description="Сохраните данные клиента, скидку и условия. Эти поля используются в расчёте и итоговых документах."
      footer={
        <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem" }}>
          <Button type="button" variant="ghost" onClick={onBack}>
            Назад
          </Button>
          <Button type="button" onClick={form.handleSubmit(onSubmit)} disabled={isPending}>
            {isPending ? "Сохраняем..." : "Далее"}
          </Button>
        </div>
      }
    >
      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      <Card title="Основные данные">
        <div style={{ display: "grid", gap: "1rem" }}>
          <FieldWrapper label="Клиент" error={form.formState.errors.clientName?.message}>
            <Input {...form.register("clientName")} placeholder="ООО Ромашка" />
          </FieldWrapper>

          <FieldWrapper label="Скидка, %" error={form.formState.errors.discountPercent?.message}>
            <Input type="number" step="0.01" min={0} max={100} {...form.register("discountPercent")} />
          </FieldWrapper>
        </div>
      </Card>

      <Card title="Условия">
        <div style={{ display: "grid", gap: "1rem" }}>
          <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap" }}>
            {([
              ["standard", "Стандартные"],
              ["custom", "Свои"],
            ] as Array<[ConditionsMode, string]>).map(([value, label]) => (
              <label
                key={value}
                style={{
                  display: "flex",
                  gap: "0.5rem",
                  alignItems: "center",
                  border: "1px solid #e4e7ec",
                  borderRadius: 12,
                  padding: "0.7rem 0.9rem",
                }}
              >
                <input type="radio" value={value} {...form.register("conditionsMode")} />
                <span>{label}</span>
              </label>
            ))}
          </div>

          {conditionsMode === "custom" && (
            <>
              <FieldWrapper label="Условия поставки" error={form.formState.errors.deliveryConditions?.message}>
                <Textarea {...form.register("deliveryConditions")} placeholder="Поставка транспортом поставщика..." />
              </FieldWrapper>
              <FieldWrapper label="Условия оплаты" error={form.formState.errors.paymentConditions?.message}>
                <Textarea {...form.register("paymentConditions")} placeholder="Оплата 50/50..." />
              </FieldWrapper>
            </>
          )}
        </div>
      </Card>
    </StepLayout>
  );
};
