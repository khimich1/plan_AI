import { useEffect, useState } from "react";
import type { CommercialDraftDetails, CommercialSaveResult, SaveMode } from "@/features/commercial-offer/types/commercialOffer";
import { DownloadFilesSection } from "@/features/commercial-offer/components/DownloadFilesSection";
import { SaveOfferSection } from "@/features/commercial-offer/components/SaveOfferSection";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { StepLayout } from "@/shared/ui/StepLayout";

type CalculationResultStepProps = {
  draft: CommercialDraftDetails;
  errorMessage: string | null;
  isGeneratingFiles: boolean;
  isSaving: boolean;
  lastSaveResult: CommercialSaveResult | null;
  executionTermsInput: string;
  onBack: () => void;
  onCreateNew: () => void;
  onGenerateFiles: () => void;
  onExecutionTermsChange: (value: string) => void;
  onSave: (payload: { mode: SaveMode; executionTermsInput: string }) => void;
  isUpdatingDiscount: boolean;
  onDiscountSubmit: (discountPercent: number) => Promise<void>;
};

export const CalculationResultStep = ({
  draft,
  errorMessage,
  isGeneratingFiles,
  isSaving,
  lastSaveResult,
  executionTermsInput,
  onBack,
  onCreateNew,
  onGenerateFiles,
  onExecutionTermsChange,
  onSave,
  isUpdatingDiscount,
  onDiscountSubmit,
}: CalculationResultStepProps) => {
  const [discountDraft, setDiscountDraft] = useState(String(draft.metadata.discount_percent ?? 0));
  const [discountError, setDiscountError] = useState<string | null>(null);
  const totalWeight = draft.order_data.reduce((acc, item) => acc + (toNumber(item.weight) ?? 0), 0);

  useEffect(() => {
    setDiscountDraft(String(draft.metadata.discount_percent ?? 0));
  }, [draft.metadata.discount_percent]);

  const handleDiscountSave = async () => {
    const parsed = toNumber(discountDraft);
    if (parsed === null || parsed < 0 || parsed > 100) {
      setDiscountError("Скидка должна быть числом от 0 до 100.");
      return;
    }
    setDiscountError(null);
    await onDiscountSubmit(parsed);
  };

  return (
    <StepLayout
    title="Шаг 5. Расчёт и результат"
    description="Запустите финальный расчёт, проверьте итоговые данные и скачайте файлы. Сохранение в БД и архив также выполняется через backend."
    footer={
      <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem" }}>
        <Button type="button" variant="ghost" onClick={onBack}>
          Назад
        </Button>
        <Button type="button" onClick={onCreateNew}>
          Создать КП
        </Button>
      </div>
    }
  >
    {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

    <Card title="Summary">
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
          gap: "0.75rem",
        }}
      >
        <SummaryCell label="Клиент" value={draft.metadata.client_name || "Не указан"} />
        <SummaryCell label="Менеджер" value={draft.metadata.manager_name || "Не выбран"} />
        <SummaryCell label="Позиций" value={String(draft.order_data.length)} />
        <SummaryCell label="Количество" value={String(draft.totals.total_qty ?? 0)} />
        <SummaryCell label="Сумма без НДС" value={`${draft.totals.subtotal ?? 0}`} />
        <SummaryCell label="Сумма с НДС" value={`${draft.totals.total_with_vat ?? 0}`} />
      </div>
    </Card>

    <Card title="Позиции">
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              {["№", "Наименование", "Кол-во", "Ед.", "Вес(кг)", "Цена", "Сумма"].map((column) => (
                <th
                  key={column}
                  style={{ textAlign: "left", padding: "0.75rem", borderBottom: "1px solid #e4e7ec" }}
                >
                  {column}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {draft.order_data.map((item, index) => (
              <tr key={`${item.name ?? "row"}-${index}`}>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>{index + 1}</td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>{String(item.name ?? "")}</td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>{String(item.qty ?? "")}</td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>шт</td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>
                  {formatNumber(item.weight)}
                </td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>
                  {formatNumber(item.unit_price)}
                </td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>
                  {formatSum(item.qty, item.unit_price)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>

    <Card title="Итоги и скидка">
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
          gap: "0.75rem",
          alignItems: "end",
        }}
      >
        <SummaryCell label="Общий вес (кг)" value={formatNumber(totalWeight)} />
        <SummaryCell label="Общая стоимость (с НДС)" value={formatNumber(draft.totals.total_with_vat ?? 0)} />
        <div style={{ border: "1px solid #e4e7ec", borderRadius: 12, padding: "0.9rem", background: "#f8fafc" }}>
          <FieldWrapper label="Скидка (%)" error={discountError}>
            <Input
              value={discountDraft}
              onChange={(event) => setDiscountDraft(event.target.value)}
              inputMode="decimal"
              placeholder="Например, 5"
            />
          </FieldWrapper>
          <Button type="button" onClick={handleDiscountSave} disabled={isUpdatingDiscount}>
            {isUpdatingDiscount ? "Обновляем..." : "Применить скидку"}
          </Button>
        </div>
      </div>
    </Card>

    <DownloadFilesSection draft={draft} isPending={isGeneratingFiles} onGenerate={onGenerateFiles} />

    <SaveOfferSection
      draft={draft}
      lastSaveResult={lastSaveResult}
      defaultExecutionTerms={executionTermsInput}
      isPending={isSaving}
      onSave={(payload) => {
        onExecutionTermsChange(payload.executionTermsInput);
        onSave(payload);
      }}
    />

    {(lastSaveResult?.result_card ?? null) && (
      <Card title="Карточка результата">
        <div style={{ display: "grid", gap: "0.5rem" }}>
          <div>Номер КП: {lastSaveResult?.result_card.offer_number}</div>
          <div>Дата: {lastSaveResult?.result_card.offer_date}</div>
          <div>Клиент: {lastSaveResult?.result_card.client_name}</div>
          <div>Менеджер: {lastSaveResult?.result_card.manager_name}</div>
          <div>Сумма: {lastSaveResult?.result_card.total_amount}</div>
          <div>Статус: {lastSaveResult?.result_card.status}</div>
          {lastSaveResult?.result_card.execution_terms && (
            <div>Срок изготовления: {lastSaveResult.result_card.execution_terms}</div>
          )}
        </div>
      </Card>
    )}
    </StepLayout>
  );
};

const SummaryCell = ({ label, value }: { label: string; value: string }) => (
  <div
    style={{
      border: "1px solid #e4e7ec",
      borderRadius: 12,
      padding: "0.9rem",
      background: "#f8fafc",
    }}
  >
    <div style={{ color: "#475467", marginBottom: "0.35rem" }}>{label}</div>
    <strong>{value}</strong>
  </div>
);

const toNumber = (value: unknown): number | null => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string" && value.trim()) {
    const parsed = Number(value.replace(",", "."));
    if (Number.isFinite(parsed)) {
      return parsed;
    }
  }
  return null;
};

const formatNumber = (value: unknown): string => {
  const parsed = toNumber(value);
  if (parsed === null) {
    return "0";
  }
  return parsed.toLocaleString("ru-RU", { maximumFractionDigits: 2 });
};

const formatSum = (qtyValue: unknown, unitPriceValue: unknown): string => {
  const qty = toNumber(qtyValue);
  const unitPrice = toNumber(unitPriceValue);
  if (qty === null || unitPrice === null) {
    return "0";
  }
  return (qty * unitPrice).toLocaleString("ru-RU", { maximumFractionDigits: 2 });
};
