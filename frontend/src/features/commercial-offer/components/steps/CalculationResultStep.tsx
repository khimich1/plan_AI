import type { CommercialDraftDetails, CommercialSaveResult, SaveMode } from "@/features/commercial-offer/types/commercialOffer";
import { DownloadFilesSection } from "@/features/commercial-offer/components/DownloadFilesSection";
import { SaveOfferSection } from "@/features/commercial-offer/components/SaveOfferSection";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
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
}: CalculationResultStepProps) => (
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
              {["Наименование", "Кол-во", "Длина", "Ширина", "Цена"].map((column) => (
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
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>{String(item.name ?? "")}</td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>{String(item.qty ?? "")}</td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>{String(item.length_m ?? "")}</td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>{String(item.width_m ?? "")}</td>
                <td style={{ padding: "0.75rem", borderBottom: "1px solid #f2f4f7" }}>{String(item.unit_price ?? "")}</td>
              </tr>
            ))}
          </tbody>
        </table>
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
