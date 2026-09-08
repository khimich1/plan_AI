import { useEffect, useMemo, useState } from "react";
import { useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { clientConditionsSchema } from "@/features/commercial-offer/schemas/commercialOffer";
import type { ConditionsMode, Manager } from "@/features/commercial-offer/types/commercialOffer";
import type { CounterpartyShort } from "@/features/commercial-offer/api/counterpartiesApi";
import {
  CounterpartyAutocomplete,
  formatCounterpartyRequisites,
} from "@/features/commercial-offer/components/CounterpartyAutocomplete";
import { useAuth } from "@/features/auth/model/AuthProvider";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Textarea } from "@/shared/ui/Field";
import { StepLayout } from "@/shared/ui/StepLayout";

type ClientConditionsDefaultValues = {
  clientName: string;
  counterpartyId: number | null;
  counterpartyCode1c?: string;
  counterpartyInn?: string | null;
  counterpartyKpp?: string | null;
  conditionsMode: ConditionsMode;
  deliveryConditions: string;
  paymentConditions: string;
};

type ClientConditionsStepProps = {
  managers: Manager[];
  selectedManagerId: number | null;
  defaultValues: ClientConditionsDefaultValues;
  errorMessage: string | null;
  isPending: boolean;
  onBack: () => void;
  onManagerChange: (managerId: number | null) => void;
  onSubmit: (payload: {
    managerId: number;
    clientName: string;
    counterpartyId: number;
    counterpartyCode1c: string;
    counterpartyInn: string | null;
    counterpartyKpp: string | null;
    conditionsMode: ConditionsMode;
    deliveryConditions: string;
    paymentConditions: string;
  }) => void;
};

type ClientConditionsFormValues = {
  clientName: string;
  counterpartyId: number;
  conditionsMode: ConditionsMode;
  deliveryConditions: string;
  paymentConditions: string;
};

const selectedFromDefaults = (values: ClientConditionsDefaultValues): CounterpartyShort | null => {
  if (values.counterpartyId == null || values.counterpartyId < 1) {
    return null;
  }
  return {
    id: values.counterpartyId,
    name: values.clientName,
    code_1c: values.counterpartyCode1c ?? "",
    inn: values.counterpartyInn ?? null,
    kpp: values.counterpartyKpp ?? null,
  };
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
  const [selected, setSelected] = useState<CounterpartyShort | null>(() => selectedFromDefaults(defaultValues));

  const form = useForm<ClientConditionsFormValues>({
    resolver: zodResolver(clientConditionsSchema),
    defaultValues: {
      ...defaultValues,
      counterpartyId: defaultValues.counterpartyId ?? 0,
    },
  });

  useEffect(() => {
    form.reset({
      ...defaultValues,
      counterpartyId: defaultValues.counterpartyId ?? 0,
    });
    setSelected(selectedFromDefaults(defaultValues));
  }, [defaultValues, form]);

  useEffect(() => {
    if (hasDefaultManager && selectedManagerId == null) {
      onManagerChange(profileManagerId);
    }
    setShowManagerOverride(!hasDefaultManager);
  }, [hasDefaultManager, onManagerChange, profileManagerId, selectedManagerId]);

  const conditionsMode = form.watch("conditionsMode");
  const counterpartyId = form.watch("counterpartyId");
  const effectiveManagerId = selectedManagerId ?? (hasDefaultManager ? profileManagerId : null);
  const selectedManager = managers.find((manager) => manager.id === effectiveManagerId) ?? null;
  const canSubmit = Boolean(effectiveManagerId) && counterpartyId > 0 && !isPending;

  const handleSelect = (item: CounterpartyShort | null) => {
    setSelected(item);
    form.setValue("counterpartyId", item?.id ?? 0, { shouldValidate: true, shouldDirty: true });
    form.setValue("clientName", item?.name ?? "", { shouldValidate: true, shouldDirty: true });
  };

  const handleSubmit = form.handleSubmit((payload) => {
    if (!effectiveManagerId || payload.counterpartyId < 1) {
      return;
    }
    onSubmit({
      managerId: effectiveManagerId,
      ...payload,
      counterpartyCode1c: selected?.code_1c ?? "",
      counterpartyInn: selected?.inn ?? null,
      counterpartyKpp: selected?.kpp ?? null,
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
          <FieldWrapper label="Клиент" error={form.formState.errors.counterpartyId?.message ?? form.formState.errors.clientName?.message}>
            <CounterpartyAutocomplete
              selected={selected}
              onSelect={handleSelect}
              placeholder="Начните вводить название или ИНН"
            />
          </FieldWrapper>
          {selected ? (
            <div
              style={{
                border: "1px solid #e4e7ec",
                borderRadius: 12,
                padding: "0.75rem 0.9rem",
                color: "#475467",
                background: "#f8fafc",
              }}
            >
              {formatCounterpartyRequisites(selected)} · код 1С {selected.code_1c || "—"}
            </div>
          ) : null}
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
