import { useEffect, useMemo, useState } from "react";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { clientConditionsSchema } from "@/features/commercial-offer/schemas/commercialOffer";
import type { ConditionsMode, Manager } from "@/features/commercial-offer/types/commercialOffer";
import { useAuth } from "@/features/auth/model/AuthProvider";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input, Textarea } from "@/shared/ui/Field";
import { StepLayout } from "@/shared/ui/StepLayout";

type ClientConditionsStepProps = {
  managers: Manager[];
  selectedManagerId: number | null;
  defaultValues: {
    clientName: string;
    conditionsMode: ConditionsMode;
    deliveryConditions: string;
    paymentConditions: string;
  };
  errorMessage: string | null;
  isPending: boolean;
  onBack: () => void;
  onManagerChange: (managerId: number | null) => void;
  onSubmit: (payload: {
    managerId: number;
    clientName: string;
    conditionsMode: ConditionsMode;
    deliveryConditions: string;
    paymentConditions: string;
  }) => void;
};

type ClientConditionsFormValues = {
  clientName: string;
  conditionsMode: ConditionsMode;
  deliveryConditions: string;
  paymentConditions: string;
};

export const ClientConditionsStep = ({
  managers,
  selectedManagerId,
  defaultValues,
  errorMessage,
  isPending,
  onBack,
  onManagerChange,
  onSubmit,
}: ClientConditionsStepProps) => {
  const { user } = useAuth();
  const profileManagerId = user?.manager_id ?? null;
  const profileManagerInList = useMemo(
    () => (profileManagerId != null ? managers.some((manager) => manager.id === profileManagerId) : false),
    [managers, profileManagerId],
  );
  const hasDefaultManager = profileManagerInList;
  const [showManagerOverride, setShowManagerOverride] = useState(!hasDefaultManager);

  const form = useForm<ClientConditionsFormValues>({
    resolver: zodResolver(clientConditionsSchema),
    defaultValues,
  });

  useEffect(() => {
    form.reset(defaultValues);
  }, [defaultValues, form]);

  useEffect(() => {
    if (hasDefaultManager && selectedManagerId == null) {
      onManagerChange(profileManagerId);
    }
    setShowManagerOverride(!hasDefaultManager);
  }, [hasDefaultManager, onManagerChange, profileManagerId, selectedManagerId]);

  const conditionsMode = form.watch("conditionsMode");
  const effectiveManagerId = selectedManagerId ?? (hasDefaultManager ? profileManagerId : null);
  const selectedManager = managers.find((manager) => manager.id === effectiveManagerId) ?? null;
  const canSubmit = Boolean(effectiveManagerId) && !isPending;

  const handleSubmit = form.handleSubmit((payload) => {
    if (!effectiveManagerId) {
      return;
    }
    onSubmit({
      managerId: effectiveManagerId,
      ...payload,
    });
  });

  return (
    <StepLayout
      title="Шаг 2. Клиент"
      description="Укажите клиента и условия. Менеджер подставляется из вашего профиля — при необходимости можно выбрать другого."
      footer={
        <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem" }}>
          <Button type="button" variant="ghost" onClick={onBack}>
            Назад
          </Button>
          <Button type="button" onClick={() => void handleSubmit()} disabled={!canSubmit}>
            {isPending ? "Рассчитываем..." : "Рассчитать КП"}
          </Button>
        </div>
      }
    >
      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      <Card title="Менеджер">
        {hasDefaultManager && selectedManager ? (
          <div style={{ display: "grid", gap: "0.75rem" }}>
            <div style={{ display: "grid", gap: "0.25rem" }}>
              <strong>{selectedManager.fio}</strong>
              <span style={{ color: "#475467" }}>{selectedManager.contact_number || "Телефон не указан"}</span>
              <span style={{ color: "#475467" }}>{selectedManager.email || "Email не указан"}</span>
            </div>
            <button
              type="button"
              onClick={() => setShowManagerOverride((open) => !open)}
              style={{
                border: "none",
                background: "none",
                color: "#175cd3",
                cursor: "pointer",
                padding: 0,
                font: "inherit",
                textAlign: "left",
              }}
            >
              {showManagerOverride ? "▾ Другой менеджер" : "▸ Другой менеджер"}
            </button>
          </div>
        ) : (
          <Alert tone="info">Выберите менеджера для итоговых документов КП.</Alert>
        )}

        {(showManagerOverride || !hasDefaultManager) && (
          <div style={{ display: "grid", gap: "0.75rem", marginTop: hasDefaultManager ? "0.75rem" : 0 }}>
            <FieldWrapper label="Менеджер">
              <select
                value={effectiveManagerId ?? ""}
                onChange={(event) => {
                  const value = event.target.value;
                  onManagerChange(value ? Number(value) : null);
                }}
                style={{
                  width: "100%",
                  border: "1px solid #d0d5dd",
                  borderRadius: 12,
                  padding: "0.8rem 0.9rem",
                  background: "#ffffff",
                }}
              >
                <option value="">Выберите менеджера</option>
                {managers.map((manager) => (
                  <option key={manager.id} value={manager.id}>
                    {manager.fio}
                    {manager.contact_number ? ` · ${manager.contact_number}` : ""}
                  </option>
                ))}
              </select>
            </FieldWrapper>
          </div>
        )}
      </Card>

      <Card title="Основные данные">
        <div style={{ display: "grid", gap: "1rem" }}>
          <FieldWrapper label="Клиент" error={form.formState.errors.clientName?.message}>
            <Input {...form.register("clientName")} placeholder="ООО Ромашка" />
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
