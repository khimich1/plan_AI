import { useEffect, useMemo, useState } from "react";
import type {
  BreakdownTable,
  CommercialDraftDetails,
  CommercialSaveResult,
  ProductType,
  SaveMode,
} from "@/features/commercial-offer/types/commercialOffer";
import { DownloadFilesSection } from "@/features/commercial-offer/components/DownloadFilesSection";
import { SaveOfferSection } from "@/features/commercial-offer/components/SaveOfferSection";
import { PlatePriceBreakdownModal } from "@/features/commercial-offer/components/PlatePriceBreakdownModal";
import { findBreakdownTable } from "@/features/commercial-offer/lib/findBreakdownTable";
import { filterCompositionWarnings } from "@/features/commercial-offer/lib/compositionWarnings";
import {
  baseProductsTotal,
  discountPercentFromTargetSum,
  formatDiscountPercentInput,
  requiresHighDiscountConfirmation,
  targetSumFromDiscountPercent,
} from "@/features/commercial-offer/lib/discountFromTargetSum";
import { HighDiscountConfirmDialog } from "@/features/commercial-offer/components/HighDiscountConfirmDialog";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import {
  formatOfferNumber,
  formatOfferSum,
  formatTotalsMoney,
  toNumber,
} from "@/features/commercial-offer/lib/formatOfferNumbers";
import { LineRowActions } from "@/features/commercial-offer/components/LineRowActions";
import { LineUndoToast } from "@/features/commercial-offer/components/LineUndoToast";
import { formatLineSourceText } from "@/features/commercial-offer/lib/formatLineSourceText";
import type { LineSavePayload, LineUndoToastState, LineRowErrorState } from "@/features/commercial-offer/lib/lineRowHandlers";
import { StepLayout } from "@/shared/ui/StepLayout";

const PRODUCT_TYPE_LABELS: Record<ProductType, string> = {
  plates: "Плиты",
  piles: "Сваи",
  steps: "Ступени",
  marches: "Марши",
  bridge_piles: "Мостовые сваи",
  fbs: "ФБС",
};

const formatProductTypeLabel = (productType: unknown): string => {
  if (typeof productType === "string" && productType in PRODUCT_TYPE_LABELS) {
    return PRODUCT_TYPE_LABELS[productType as ProductType];
  }
  return typeof productType === "string" && productType.length > 0 ? productType : "—";
};

const thStyle = { textAlign: "left" as const, padding: "0.75rem", borderBottom: "1px solid #e4e7ec" };
const tdStyle = { padding: "0.75rem", borderBottom: "1px solid #f2f4f7" };

type CalculationResultStepProps = {
  draft: CommercialDraftDetails;
  isPileDraft?: boolean;
  isStepDraft?: boolean;
  isMarchDraft?: boolean;
  isBridgePileDraft?: boolean;
  isFbsDraft?: boolean;
  isSimpleKpDraft?: boolean;
  breakdownTables: BreakdownTable[];
  isBreakdownLoading: boolean;
  errorMessage: string | null;
  isGeneratingFiles: boolean;
  isGeneratingSchema: boolean;
  isSaving: boolean;
  lastSaveResult: CommercialSaveResult | null;
  executionTermsInput: string;
  onBack: () => void;
  onCreateNew: () => void;
  onGenerateFiles: () => void;
  onGenerateSchema: () => void;
  onExecutionTermsChange: (value: string) => void;
  onSave: (payload: { mode: SaveMode; executionTermsInput: string }) => Promise<void>;
  isUpdatingDiscount: boolean;
  onDiscountSubmit: (discountPercent: number) => Promise<void>;
  onLogisticsCostSubmit: (logisticsCost: number) => Promise<void>;
  onAddOtherNomenclature?: () => void;
  onUndoLastBatch?: () => Promise<void> | void;
  onDeleteLine?: (lineId: string) => Promise<void> | void;
  onSaveLine?: (lineId: string, payload: LineSavePayload) => Promise<void> | void;
  lineUndoToast?: LineUndoToastState | null;
  lineRowError?: LineRowErrorState | null;
};

export const CalculationResultStep = ({
  draft,
  isPileDraft = false,
  isStepDraft = false,
  isMarchDraft = false,
  isBridgePileDraft = false,
  isFbsDraft = false,
  isSimpleKpDraft = false,
  breakdownTables,
  isBreakdownLoading,
  errorMessage,
  isGeneratingFiles,
  isGeneratingSchema,
  isSaving,
  lastSaveResult,
  executionTermsInput,
  onBack,
  onCreateNew,
  onGenerateFiles,
  onGenerateSchema,
  onExecutionTermsChange,
  onSave,
  isUpdatingDiscount,
  onDiscountSubmit,
  onLogisticsCostSubmit,
  onAddOtherNomenclature,
  onUndoLastBatch,
  onDeleteLine,
  onSaveLine,
  lineUndoToast = null,
  lineRowError = null,
}: CalculationResultStepProps) => {
  const [discountDraft, setDiscountDraft] = useState(String(draft.metadata.discount_percent ?? 0));
  const [targetSumDraft, setTargetSumDraft] = useState("");
  const [logisticsCostDraft, setLogisticsCostDraft] = useState(String(draft.metadata.logistics_cost ?? 0));
  const [discountError, setDiscountError] = useState<string | null>(null);
  const [targetSumError, setTargetSumError] = useState<string | null>(null);
  const [logisticsError, setLogisticsError] = useState<string | null>(null);
  const [selectedPlateName, setSelectedPlateName] = useState<string | null>(null);
  const [pendingDiscountPercent, setPendingDiscountPercent] = useState<number | null>(null);
  const isGradeSimpleDraft = isPileDraft || isMarchDraft || isBridgePileDraft || isFbsDraft;
  const breakdownAvailable = !isSimpleKpDraft && (draft.metadata.breakdown_tables_count ?? 0) > 0;
  const selectedBreakdownTable = useMemo(
    () => (selectedPlateName ? findBreakdownTable(breakdownTables, selectedPlateName) : undefined),
    [breakdownTables, selectedPlateName],
  );
  const appendBatches = draft.metadata.append_batches ?? [];
  const distinctProductTypes = useMemo(() => {
    const types = new Set<string>();
    for (const item of draft.order_data) {
      if (typeof item.product_type === "string" && item.product_type.length > 0) {
        types.add(item.product_type);
      }
    }
    return types;
  }, [draft.order_data]);
  const showTypeColumn = distinctProductTypes.size > 1 || appendBatches.length > 1;
  const hasPlateLines = draft.order_data.some((item) => item.product_type === "plates");
  const tripCostDisabled = !hasPlateLines;
  const totalWeight = draft.order_data.reduce((acc, item) => acc + (toNumber(item.weight) ?? 0), 0);
  const serverSubtotal = draft.totals.subtotal;
  const serverVat = draft.totals.vat_amount;
  const serverTotalWithVat = draft.totals.total_with_vat;
  const baseProducts = useMemo(() => baseProductsTotal(draft.order_data), [draft.order_data]);
  const savedDiscountPercent = draft.metadata.discount_percent ?? 0;
  const derivedDelivery = useMemo(() => {
    if (typeof serverTotalWithVat !== "number" || !Number.isFinite(serverTotalWithVat) || baseProducts <= 0) {
      return null;
    }
    const delivery = serverTotalWithVat - baseProducts * (1 - savedDiscountPercent / 100);
    return delivery >= -0.01 ? Math.max(0, delivery) : null;
  }, [baseProducts, savedDiscountPercent, serverTotalWithVat]);
  const savedTargetSum =
    derivedDelivery === null
      ? null
      : targetSumFromDiscountPercent({
          discountPercent: savedDiscountPercent,
          baseProductsTotalWithVat: baseProducts,
          deliveryTotal: derivedDelivery,
        });

  useEffect(() => {
    setDiscountDraft(String(draft.metadata.discount_percent ?? 0));
    setTargetSumDraft(savedTargetSum === null ? "" : String(savedTargetSum).replace(".", ","));
    setDiscountError(null);
    setTargetSumError(null);
  }, [draft.metadata.discount_percent, savedTargetSum]);

  useEffect(() => {
    setLogisticsCostDraft(String(draft.metadata.logistics_cost ?? 0).replace(".", ","));
  }, [draft.metadata.logistics_cost]);

  const restoreDiscountDrafts = () => {
    setDiscountDraft(String(savedDiscountPercent));
    setTargetSumDraft(savedTargetSum === null ? "" : String(savedTargetSum).replace(".", ","));
    setDiscountError(null);
    setTargetSumError(null);
  };

  const applyDiscount = async (discountPercent: number) => {
    setPendingDiscountPercent(null);
    await onDiscountSubmit(discountPercent);
  };

  const requestDiscountApply = (discountPercent: number) => {
    if (requiresHighDiscountConfirmation(discountPercent)) {
      setPendingDiscountPercent(discountPercent);
      return;
    }
    void applyDiscount(discountPercent);
  };

  const handleDiscountSave = () => {
    const parsed = toNumber(discountDraft);
    if (parsed === null || parsed < 0 || parsed > 100) {
      setDiscountError("Скидка должна быть числом от 0 до 100.");
      return;
    }
    setDiscountError(null);
    if (derivedDelivery === null || baseProducts <= 0) {
      setDiscountError("Не удалось определить стоимость доставки для расчёта целевой суммы.");
      return;
    }
    const target = targetSumFromDiscountPercent({
      discountPercent: parsed,
      baseProductsTotalWithVat: baseProducts,
      deliveryTotal: derivedDelivery,
    });
    if (target === null) {
      setDiscountError("Скидка должна быть числом от 0 до 100.");
      return;
    }
    setTargetSumDraft(String(target).replace(".", ","));
    requestDiscountApply(parsed);
  };

  const handleTargetSumChange = (value: string) => {
    setTargetSumDraft(value);
    const parsed = toNumber(value);
    if (parsed === null) {
      setTargetSumError(value.trim() ? "Введите корректную целевую сумму." : null);
      return;
    }
    if (derivedDelivery === null) {
      setTargetSumError("Не удалось определить стоимость доставки для расчёта целевой суммы.");
      return;
    }
    const result = discountPercentFromTargetSum({
      targetTotalWithVat: parsed,
      baseProductsTotalWithVat: baseProducts,
      deliveryTotal: derivedDelivery,
    });
    if (!result.ok) {
      setTargetSumError(result.error);
      return;
    }
    setTargetSumError(null);
    setDiscountError(null);
    setDiscountDraft(formatDiscountPercentInput(result.discountPercent));
  };

  const handleTargetSumSave = () => {
    const parsed = toNumber(targetSumDraft);
    if (parsed === null || derivedDelivery === null) {
      setTargetSumError("Введите корректную целевую сумму.");
      return;
    }
    const result = discountPercentFromTargetSum({
      targetTotalWithVat: parsed,
      baseProductsTotalWithVat: baseProducts,
      deliveryTotal: derivedDelivery,
    });
    if (!result.ok) {
      setTargetSumError(result.error);
      return;
    }
    setTargetSumError(null);
    setDiscountDraft(formatDiscountPercentInput(result.discountPercent));
    requestDiscountApply(result.discountPercent);
  };

  const handleApplyLogisticsCost = async () => {
    const parsed = toNumber(logisticsCostDraft);
    if (parsed === null || parsed < 0) {
      setLogisticsError("Стоимость рейса должна быть числом не меньше 0.");
      return;
    }
    setLogisticsError(null);
    await onLogisticsCostSubmit(parsed);
    setLogisticsCostDraft(String(parsed).replace(".", ","));
  };

  const totalWithVat = formatTotalsMoney(serverTotalWithVat);
  const readinessWarnings = filterCompositionWarnings(draft.metadata.warnings);

  return (
    <StepLayout
    title="Шаг 3. Результат"
    description="Проверьте готовность КП, скачайте файлы и сохраните результат."
    footer={
      <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem" }}>
        <Button type="button" variant="ghost" onClick={onBack}>
          Назад
        </Button>
        <Button type="button" variant="danger" onClick={onCreateNew}>
          Создать новое КП
        </Button>
      </div>
    }
  >
    {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

    <Card title="Готовность КП" subtitle="Перед отправкой клиенту проверьте ключевые пункты.">
      <ul style={{ margin: 0, paddingLeft: "1.25rem", display: "grid", gap: "0.5rem" }}>
        <li>✓ {draft.order_data.length} позиций в заказе</li>
        <li>✓ {draft.totals.total_qty ?? 0} {isStepDraft ? "ступеней" : isMarchDraft ? "маршей" : isBridgePileDraft ? "мостовых свай" : isFbsDraft ? "ФБС" : isPileDraft ? "свай" : "плит"} в заказе</li>
        <li>✓ Клиент: {draft.metadata.client_name || "не указан"}</li>
        <li>✓ Сумма с НДС: {totalWithVat}</li>
        {readinessWarnings.length > 0 && (
          <li style={{ color: "#b54708" }}>
            ⚠ {readinessWarnings.length} предупреждени{readinessWarnings.length === 1 ? "е" : readinessWarnings.length < 5 ? "я" : "й"}:
            <ul style={{ margin: "0.35rem 0 0", paddingLeft: "1.25rem" }}>
              {readinessWarnings.slice(0, 5).map((warning) => (
                <li key={warning}>{warning}</li>
              ))}
              {readinessWarnings.length > 5 && <li>… и ещё {readinessWarnings.length - 5}</li>}
            </ul>
          </li>
        )}
      </ul>
    </Card>

    <Card title="Сводка">
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
        <SummaryCell label="Сумма без НДС" value={formatTotalsMoney(serverSubtotal)} />
        <SummaryCell label="НДС" value={formatTotalsMoney(serverVat)} />
        <SummaryCell label="Сумма с НДС" value={formatTotalsMoney(serverTotalWithVat)} />
      </div>
    </Card>

    <Card title="Позиции">
      <div style={{ display: "grid", gap: "0.75rem" }}>
        {lineUndoToast ? (
          <LineUndoToast message={lineUndoToast.message} onUndo={lineUndoToast.onUndo} />
        ) : null}
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr>
              {(isStepDraft
                ? ["№", "Марка", "Кол-во", "Цена", "Сумма"]
                : isGradeSimpleDraft
                  ? ["№", "Марка", "Класс", "Кол-во", "Цена", "Сумма"]
                  : ["№", "Наименование", "Кол-во", "Ед.", "Вес(кг)", "Цена", "Сумма"]
              )
                .flatMap((column, columnIndex) =>
                  columnIndex === 1 && showTypeColumn ? ["Тип", column] : [column],
                )
                .concat([""])
                .map((column, columnIndex) => (
                  <th key={`${column || "actions"}-${columnIndex}`} style={thStyle}>
                    {column}
                  </th>
                ))}
            </tr>
          </thead>
          <tbody>
            {draft.order_data.map((item, index) => {
              const itemName = String(item.name ?? item.mark ?? "");
              const canOpenBreakdown = breakdownAvailable && !isBreakdownLoading && itemName.length > 0;
              const lineId = typeof item.line_id === "string" ? item.line_id : null;
              const typeCell = showTypeColumn ? (
                <td style={tdStyle}>{formatProductTypeLabel(item.product_type)}</td>
              ) : null;
              const actionCell = (
                <td style={tdStyle}>
                  {lineId ? (
                    <LineRowActions
                      lineId={lineId}
                      defaultQty={toNumber(item.qty) ?? 0}
                      defaultSourceText={formatLineSourceText(item)}
                      saveError={lineRowError?.lineId === lineId ? lineRowError.message : null}
                      onSave={(payload) => void onSaveLine?.(lineId, payload)}
                      onDelete={() => void onDeleteLine?.(lineId)}
                    />
                  ) : null}
                </td>
              );

              if (isStepDraft) {
                return (
                  <tr key={lineId ?? `${itemName}-${index}`}>
                    <td style={tdStyle}>{index + 1}</td>
                    {typeCell}
                    <td style={tdStyle}>{itemName}</td>
                    <td style={tdStyle}>{String(item.qty ?? "")}</td>
                    <td style={tdStyle}>{formatOfferNumber(item.unit_price)}</td>
                    <td style={tdStyle}>{formatOfferSum(item.qty, item.unit_price)}</td>
                    {actionCell}
                  </tr>
                );
              }

              if (isGradeSimpleDraft) {
                return (
                  <tr key={lineId ?? `${itemName}-${index}`}>
                    <td style={tdStyle}>{index + 1}</td>
                    {typeCell}
                    <td style={tdStyle}>{itemName}</td>
                    <td style={tdStyle}>{String(item.concrete_grade ?? "—")}</td>
                    <td style={tdStyle}>{String(item.qty ?? "")}</td>
                    <td style={tdStyle}>{formatOfferNumber(item.unit_price)}</td>
                    <td style={tdStyle}>{formatOfferSum(item.qty, item.unit_price)}</td>
                    {actionCell}
                  </tr>
                );
              }

              const plateName = itemName;
              return (
                <tr key={lineId ?? `${item.name ?? "row"}-${index}`}>
                  <td style={tdStyle}>{index + 1}</td>
                  {typeCell}
                  <td style={tdStyle}>
                    {canOpenBreakdown ? (
                      <button
                        type="button"
                        onClick={() => setSelectedPlateName(plateName)}
                        title="Показать детальную разбивку цены"
                        style={{
                          color: "#175cd3",
                          textDecoration: "underline",
                          cursor: "pointer",
                          background: "none",
                          border: "none",
                          padding: 0,
                          textAlign: "left",
                          font: "inherit",
                        }}
                      >
                        {plateName}
                      </button>
                    ) : (
                      plateName
                    )}
                  </td>
                  <td style={tdStyle}>{String(item.qty ?? "")}</td>
                  <td style={tdStyle}>шт</td>
                  <td style={tdStyle}>{formatOfferNumber(item.weight)}</td>
                  <td style={tdStyle}>{formatOfferNumber(item.unit_price)}</td>
                  <td style={tdStyle}>{formatOfferSum(item.qty, item.unit_price)}</td>
                  {actionCell}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "0.75rem", marginTop: "1rem" }}>
        <Button type="button" variant="secondary" onClick={() => onAddOtherNomenclature?.()}>
          Добавить другое наименование
        </Button>
        {appendBatches.length > 0 && (
          <Button type="button" variant="ghost" onClick={() => void onUndoLastBatch?.()}>
            Отменить последний заход
          </Button>
        )}
      </div>
    </Card>

    <Card title="Итоги и скидка">
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
          gap: "0.75rem",
          alignItems: "start",
        }}
      >
        <div style={{ display: "grid", gap: "0.75rem" }}>
          {!isSimpleKpDraft && <SummaryCell label="Общий вес (кг)" value={formatOfferNumber(totalWeight)} />}
          <div style={{ border: "1px solid #e4e7ec", borderRadius: 12, padding: "0.9rem", background: "#f8fafc" }}>
            <FieldWrapper label="Стоимость рейса" error={logisticsError}>
              <div style={{ position: "relative" }}>
                <input
                  value={logisticsCostDraft}
                  onChange={(event) => setLogisticsCostDraft(event.target.value)}
                  inputMode="decimal"
                  placeholder="Стоимость одного рейса"
                  disabled={tripCostDisabled}
                  style={{
                    width: "100%",
                    border: "1px solid #d0d5dd",
                    borderRadius: 12,
                    padding: "0.8rem 3.5rem 0.8rem 0.9rem",
                    background: tripCostDisabled ? "#f2f4f7" : "#ffffff",
                  }}
                />
                <Button
                  type="button"
                  variant="secondary"
                  onClick={handleApplyLogisticsCost}
                  disabled={isUpdatingDiscount || tripCostDisabled}
                  style={{
                    position: "absolute",
                    right: "0.35rem",
                    top: "50%",
                    transform: "translateY(-50%)",
                    borderRadius: 8,
                    padding: "0.25rem 0.6rem",
                    fontSize: "0.8rem",
                  }}
                >
                  OK
                </Button>
              </div>
            </FieldWrapper>
          </div>
        </div>
        <div style={{ display: "grid", gap: "0.75rem" }}>
          <SummaryCell label="НДС" value={formatTotalsMoney(serverVat)} />
          <SummaryCell label="Стоимость с НДС" value={formatTotalsMoney(serverTotalWithVat)} />
        </div>
        <div style={{ border: "1px solid #e4e7ec", borderRadius: 12, padding: "0.9rem", background: "#f8fafc" }}>
          <FieldWrapper label="Целевая сумма (₽)" error={targetSumError}>
            <Input
              value={targetSumDraft}
              onChange={(event) => handleTargetSumChange(event.target.value)}
              inputMode="decimal"
              placeholder="Например, 2 000 000"
              disabled={baseProducts <= 0 || derivedDelivery === null}
            />
          </FieldWrapper>
          <Button type="button" variant="secondary" onClick={handleTargetSumSave} disabled={isUpdatingDiscount || baseProducts <= 0 || derivedDelivery === null}>
            Применить сумму
          </Button>
          <FieldWrapper label="Скидка (%)" error={discountError}>
            <Input
              value={discountDraft}
              onChange={(event) => {
                setDiscountDraft(event.target.value);
                const discount = toNumber(event.target.value);
                if (discount === null || derivedDelivery === null) {
                  return;
                }
                const target = targetSumFromDiscountPercent({
                  discountPercent: discount,
                  baseProductsTotalWithVat: baseProducts,
                  deliveryTotal: derivedDelivery,
                });
                if (target !== null) {
                  setTargetSumDraft(String(target).replace(".", ","));
                  setTargetSumError(null);
                }
              }}
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

    <DownloadFilesSection
      draft={draft}
      isSimpleKpDraft={isSimpleKpDraft}
      isPending={isGeneratingFiles}
      isSchemaPending={isGeneratingSchema}
      onGenerate={onGenerateFiles}
      onGenerateSchema={onGenerateSchema}
    />

    <SaveOfferSection
      draft={draft}
      lastSaveResult={lastSaveResult}
      defaultExecutionTerms={executionTermsInput}
      isPending={isSaving}
      onSave={async (payload) => {
        onExecutionTermsChange(payload.executionTermsInput);
        await onSave(payload);
      }}
    />

    {!isSimpleKpDraft && (
      <PlatePriceBreakdownModal
        open={selectedPlateName !== null}
        plateName={selectedPlateName}
        table={selectedBreakdownTable}
        onClose={() => setSelectedPlateName(null)}
      />
    )}

    <HighDiscountConfirmDialog
      open={pendingDiscountPercent !== null}
      discountPercent={pendingDiscountPercent ?? 0}
      isPending={isUpdatingDiscount}
      onConfirm={() => pendingDiscountPercent !== null && void applyDiscount(pendingDiscountPercent)}
      onCancel={() => {
        setPendingDiscountPercent(null);
        restoreDiscountDrafts();
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
