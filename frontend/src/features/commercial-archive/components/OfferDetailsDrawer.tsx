import { useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";
import { Alert } from "@/shared/ui/Alert";
import { FieldWrapper } from "@/shared/ui/Field";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import {
  useArchiveOfferQuery,
  useArchiveDocumentMutation,
  useUpdateDiscountMutation,
  useUpdateLogisticsCostMutation,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import {
  cargoDeliveryTripsCount,
} from "@/features/commercial-offer/utils/cargoDeliveryPricing";
import { formatMoney, statusEmoji } from "@/features/commercial-archive/lib/format";
import type { ArchiveOfferDetails, ArchiveMarchItem, ArchivePileItem, ArchiveStepItem } from "@/features/commercial-archive/types/archive";
import { useWizardDraftStore } from "@/features/commercial-offer/store/wizardDraftStore";
import { DeliveryScheduleDialog } from "@/features/delivery-schedule/components/DeliveryScheduleDialog";
import { useDeliveryScheduleQuery } from "@/features/delivery-schedule/hooks/useDeliveryScheduleQueries";
import type { OfferPlateForSchedule } from "@/features/delivery-schedule/lib/scheduleDraft";
import { downloadFile } from "@/shared/lib/downloadFile";
import { getErrorMessage } from "@/shared/lib/apiError";
import { DeleteConfirmDialog } from "./DeleteConfirmDialog";
import { MoveToProductionDialog } from "./MoveToProductionDialog";
import { KpReadinessBlock } from "./KpReadinessBlock";
import { HighDiscountConfirmDialog } from "@/features/commercial-offer/components/HighDiscountConfirmDialog";
import {
  baseProductsTotal,
  discountPercentFromTargetSum,
  formatDiscountPercentInput,
  requiresHighDiscountConfirmation,
  targetSumFromDiscountPercent,
} from "@/features/commercial-offer/lib/discountFromTargetSum";

const DELIVERY_SCHEDULE_EDITABLE_STATUSES = new Set(["в работе", "На СГП"]);
/** График поставки недоступен в секции «В архиве» — только после перевода в производство. */
const DELIVERY_SCHEDULE_HIDDEN_STATUSES = new Set(["в архиве"]);

type Props = {
  open: boolean;
  kpId: number | null;
  onClose: () => void;
};

const PLATES_PREVIEW = 10;

const resolvePileItems = (offer: ArchiveOfferDetails): ArchivePileItem[] => {
  if (offer.piles && offer.piles.length > 0) {
    return offer.piles;
  }
  return offer.plates.map((plate, index) => ({
    position_number: plate.position_number ?? index + 1,
    mark: plate.plate_name,
    concrete_grade: "—",
    qty: plate.qty,
    unit_price: plate.unit_price,
    discounted_price: plate.discounted_price,
  }));
};

const resolveStepItems = (offer: ArchiveOfferDetails): ArchiveStepItem[] => offer.steps ?? [];

const resolveBridgePileItems = (offer: ArchiveOfferDetails): ArchivePileItem[] => {
  if (offer.bridge_piles && offer.bridge_piles.length > 0) {
    return offer.bridge_piles;
  }
  return [];
};

const resolveFbsItems = (offer: ArchiveOfferDetails): ArchivePileItem[] => {
  if (offer.fbs && offer.fbs.length > 0) {
    return offer.fbs;
  }
  return [];
};

const resolveMarchItems = (offer: ArchiveOfferDetails): ArchiveMarchItem[] => {
  if (offer.marches && offer.marches.length > 0) {
    return offer.marches;
  }
  return offer.plates.map((plate, index) => ({
    position_number: plate.position_number ?? index + 1,
    mark: plate.plate_name,
    concrete_grade: "—",
    qty: plate.qty,
    unit_price: plate.unit_price,
    discounted_price: plate.discounted_price,
  }));
};

const formatLinePrice = (discountedPrice: number | null, unitPrice: number | null): string => {
  if (discountedPrice !== null) {
    return formatMoney(discountedPrice);
  }
  if (unitPrice !== null) {
    return formatMoney(unitPrice);
  }
  return "—";
};

export const OfferDetailsDrawer = ({ open, kpId, onClose }: Props) => {
  const navigate = useNavigate();
  const { dispatch } = useWizardDraftStore();
  const query = useArchiveOfferQuery(open ? kpId : null);
  const [showAllPlates, setShowAllPlates] = useState(false);
  const [moveOpen, setMoveOpen] = useState(false);
  const [discountDraft, setDiscountDraft] = useState("");
  const [targetSumDraft, setTargetSumDraft] = useState("");
  const [logisticsDraft, setLogisticsDraft] = useState("");
  const [discountError, setDiscountError] = useState<string | null>(null);
  const [targetSumError, setTargetSumError] = useState<string | null>(null);
  const [logisticsError, setLogisticsError] = useState<string | null>(null);
  const [deleteOpen, setDeleteOpen] = useState(false);
  const [pendingDiscountPercent, setPendingDiscountPercent] = useState<number | null>(null);
  const [resumePending, setResumePending] = useState(false);
  const [resumeError, setResumeError] = useState<string | null>(null);

  const discountMutation = useUpdateDiscountMutation();
  const logisticsMutation = useUpdateLogisticsCostMutation();
  const schemaMutation = useArchiveDocumentMutation("schema");
  const financePending = discountMutation.isPending || logisticsMutation.isPending;
  const [scheduleOpen, setScheduleOpen] = useState(false);

  const offer = query.data;
  const isPileOffer = offer?.product_type === "piles";
  const isStepOffer = offer?.product_type === "steps";
  const isMarchOffer = offer?.product_type === "marches";
  const isBridgePileOffer = offer?.product_type === "bridge_piles";
  const isFbsOffer = offer?.product_type === "fbs";
  const isSimpleProductOffer = isPileOffer || isStepOffer || isMarchOffer || isBridgePileOffer || isFbsOffer;
  const showReadiness =
    !isSimpleProductOffer && (offer?.status === "в работе" || offer?.status === "На СГП");
  const canShowDeliverySchedule =
    !isSimpleProductOffer &&
    offer?.status != null &&
    !DELIVERY_SCHEDULE_HIDDEN_STATUSES.has(offer.status);
  const canEditDeliverySchedule =
    Boolean(
      canShowDeliverySchedule &&
        offer?.status &&
        DELIVERY_SCHEDULE_EDITABLE_STATUSES.has(offer.status),
    );
  const scheduleQuery = useDeliveryScheduleQuery(
    open && offer && canShowDeliverySchedule ? offer.kp_id : null,
  );
  const hasDeliverySchedule = scheduleQuery.data != null;
  const schedulePlates: OfferPlateForSchedule[] = useMemo(() => {
    if (!offer || isSimpleProductOffer) {
      return [];
    }
    return offer.plates
      .filter((plate): plate is typeof plate & { id: number } => typeof plate.id === "number" && plate.id > 0)
      .map((plate) => ({
        id: plate.id,
        plate_name: plate.plate_name,
        qty: plate.qty,
        position_number: plate.position_number,
      }));
  }, [offer, isSimpleProductOffer]);
  const pileItems = offer && isPileOffer ? resolvePileItems(offer) : [];
  const stepItems = offer && isStepOffer ? resolveStepItems(offer) : [];
  const marchItems = offer && isMarchOffer ? resolveMarchItems(offer) : [];
  const bridgePileItems = offer && isBridgePileOffer ? resolveBridgePileItems(offer) : [];
  const fbsItems = offer && isFbsOffer ? resolveFbsItems(offer) : [];
  const allProductItems = offer
    ? isPileOffer
      ? pileItems
      : isStepOffer
        ? stepItems
        : isMarchOffer
          ? marchItems
          : isBridgePileOffer
            ? bridgePileItems
            : isFbsOffer
              ? fbsItems
              : offer.plates
    : [];
  const baseProducts = baseProductsTotal(allProductItems as unknown as Array<Record<string, unknown>>);
  const deliveryTotal = offer?.delivery_service_total_rub ?? 0;
  const savedTargetSum =
    offer && baseProducts > 0
      ? targetSumFromDiscountPercent({
          discountPercent: offer.finance.discount_percent,
          baseProductsTotalWithVat: baseProducts,
          deliveryTotal,
        })
      : null;
  const orderItemsCount = isPileOffer
    ? pileItems.length
    : isStepOffer
      ? stepItems.length
      : isMarchOffer
        ? marchItems.length
        : isBridgePileOffer
          ? bridgePileItems.length
          : isFbsOffer
            ? fbsItems.length
          : (offer?.plates.length ?? 0);
  const itemsToShow = offer
    ? showAllPlates
      ? isPileOffer
        ? pileItems
        : isStepOffer
          ? stepItems
          : isMarchOffer
            ? marchItems
            : isBridgePileOffer
              ? bridgePileItems
              : isFbsOffer
                ? fbsItems
                : offer.plates
      : isPileOffer
        ? pileItems.slice(0, PLATES_PREVIEW)
        : isStepOffer
          ? stepItems.slice(0, PLATES_PREVIEW)
          : isMarchOffer
            ? marchItems.slice(0, PLATES_PREVIEW)
            : isBridgePileOffer
              ? bridgePileItems.slice(0, PLATES_PREVIEW)
              : isFbsOffer
                ? fbsItems.slice(0, PLATES_PREVIEW)
                : offer.plates.slice(0, PLATES_PREVIEW)
    : [];

  useEffect(() => {
    if (offer) {
      setDiscountDraft(String(offer.finance.discount_percent ?? 0));
      setTargetSumDraft(savedTargetSum === null ? "" : String(savedTargetSum).replace(".", ","));
      setLogisticsDraft(String(offer.logistics_cost ?? 0).replace(".", ","));
      setDiscountError(null);
      setTargetSumError(null);
      setLogisticsError(null);
      setResumeError(null);
    }
  }, [offer?.kp_id, offer?.finance.discount_percent, offer?.logistics_cost, offer?.delivery_service_total_rub, savedTargetSum]);

  const restoreDiscountDrafts = () => {
    if (!offer) {
      return;
    }
    setDiscountDraft(String(offer.finance.discount_percent));
    setTargetSumDraft(savedTargetSum === null ? "" : String(savedTargetSum).replace(".", ","));
    setDiscountError(null);
    setTargetSumError(null);
  };

  const handleResumeAppend = async () => {
    if (!offer || resumePending) {
      return;
    }
    setResumePending(true);
    setResumeError(null);
    try {
      const draft = await archiveApi.resume(offer.kp_id);
      dispatch({ type: "hydrate-draft", payload: draft });
      dispatch({ type: "start-append-cycle" });
      navigate(`/new?draft=${encodeURIComponent(draft.draft_id)}`);
      onClose();
    } catch (error) {
      setResumeError(getErrorMessage(error));
    } finally {
      setResumePending(false);
    }
  };

  const applyDiscount = async (discount: number) => {
    if (!offer) {
      return;
    }
    setPendingDiscountPercent(null);
    await discountMutation.mutateAsync({ kpId: offer.kp_id, discount });
  };

  const requestDiscountApply = (discount: number) => {
    if (requiresHighDiscountConfirmation(discount)) {
      setPendingDiscountPercent(discount);
      return;
    }
    void applyDiscount(discount);
  };

  const handleDiscountOk = () => {
    if (!offer) {
      return;
    }
    const parsed = parseNumberField(discountDraft);
    if (parsed === null || parsed < 0 || parsed > 100) {
      setDiscountError("Скидка должна быть числом от 0 до 100.");
      return;
    }
    setDiscountError(null);
    const target = targetSumFromDiscountPercent({
      discountPercent: parsed,
      baseProductsTotalWithVat: baseProducts,
      deliveryTotal,
    });
    if (target === null) {
      setDiscountError("Нет позиций для расчёта скидки.");
      return;
    }
    setTargetSumDraft(String(target).replace(".", ","));
    requestDiscountApply(parsed);
  };

  const handleTargetSumChange = (value: string) => {
    setTargetSumDraft(value);
    const target = parseNumberField(value);
    if (target === null) {
      setTargetSumError(value.trim() ? "Введите корректную целевую сумму." : null);
      return;
    }
    const result = discountPercentFromTargetSum({
      targetTotalWithVat: target,
      baseProductsTotalWithVat: baseProducts,
      deliveryTotal,
    });
    if (!result.ok) {
      setTargetSumError(result.error);
      return;
    }
    setTargetSumError(null);
    setDiscountError(null);
    setDiscountDraft(formatDiscountPercentInput(result.discountPercent));
  };

  const handleTargetSumOk = () => {
    const target = parseNumberField(targetSumDraft);
    if (target === null) {
      setTargetSumError("Введите корректную целевую сумму.");
      return;
    }
    const result = discountPercentFromTargetSum({
      targetTotalWithVat: target,
      baseProductsTotalWithVat: baseProducts,
      deliveryTotal,
    });
    if (!result.ok) {
      setTargetSumError(result.error);
      return;
    }
    setTargetSumError(null);
    setDiscountDraft(formatDiscountPercentInput(result.discountPercent));
    requestDiscountApply(result.discountPercent);
  };

  const handleLogisticsOk = async () => {
    if (!offer) {
      return;
    }
    const parsed = parseNumberField(logisticsDraft);
    if (parsed === null || parsed < 0) {
      setLogisticsError("Стоимость рейса должна быть числом не меньше 0.");
      return;
    }
    setLogisticsError(null);
    await logisticsMutation.mutateAsync({ kpId: offer.kp_id, logisticsCost: parsed });
    setLogisticsDraft(String(parsed).replace(".", ","));
  };

  const clientTrips = offer ? cargoDeliveryTripsCount(Math.max(0, offer.total_cargo_weight_kg ?? 0)) : 0;

  return (
    <Modal
      open={open}
      onClose={onClose}
      title={
        offer ? (
          <span style={{ display: "inline-flex", alignItems: "center", gap: "0.5rem", flexWrap: "wrap" }}>
            <span>КП №{offer.kp_id}</span>
            {canShowDeliverySchedule && hasDeliverySchedule && (
              <span
                style={{
                  fontSize: "0.75rem",
                  fontWeight: 600,
                  padding: "0.15rem 0.5rem",
                  borderRadius: 999,
                  background: "#ecfdf3",
                  color: "#067647",
                }}
              >
                есть график
              </span>
            )}
          </span>
        ) : (
          "Карточка КП"
        )
      }
      maxWidth={720}
    >
      {query.isPending && (
        <div style={{ display: "flex", gap: "0.75rem", alignItems: "center" }}>
          <Spinner /> Загружаю данные...
        </div>
      )}

      {query.isError && <Alert tone="error">{getErrorMessage(query.error)}</Alert>}

      {offer && (
        <div style={{ display: "grid", gap: "1rem" }}>
          <section
            style={{
              display: "grid",
              gap: "0.45rem",
              padding: "1rem",
              background: "#f8faff",
              border: "1px solid #e4e7ec",
              borderRadius: 14,
            }}
          >
            <div style={{ display: "flex", flexWrap: "wrap", gap: "0.75rem", justifyContent: "space-between" }}>
              <div>
                <div style={{ color: "#667085", fontSize: "0.85rem" }}>Клиент</div>
                <div style={{ fontWeight: 600 }}>{offer.customer_name || "—"}</div>
              </div>
              <div>
                <div style={{ color: "#667085", fontSize: "0.85rem" }}>Менеджер</div>
                <div style={{ fontWeight: 600 }}>{offer.manager_name || "—"}</div>
              </div>
              <div>
                <div style={{ color: "#667085", fontSize: "0.85rem" }}>Дата создания</div>
                <div style={{ fontWeight: 600 }}>{offer.creation_date || "—"}</div>
              </div>
              <div>
                <div style={{ color: "#667085", fontSize: "0.85rem" }}>Статус</div>
                <div style={{ fontWeight: 600 }}>
                  {statusEmoji(offer.status)} {offer.status || "—"}
                </div>
              </div>
              {offer.execution_terms && (
                <div>
                  <div style={{ color: "#667085", fontSize: "0.85rem" }}>Срок</div>
                  <div style={{ fontWeight: 600 }}>⏰ {offer.execution_terms}</div>
                </div>
              )}
            </div>
          </section>

          {showReadiness && offer.readiness && (
            <KpReadinessBlock kpId={offer.kp_id} readiness={offer.readiness} />
          )}

          {/* Итоги: слева вес, НДС, рейс, доставка; справа скидка и под ней «Итого с НДС» */}
          <section
            style={{
              padding: "1rem",
              border: "1px solid #e4e7ec",
              borderRadius: 14,
              background: "#ffffff",
            }}
          >
            <h3 style={{ margin: "0 0 0.75rem", fontSize: "1rem" }}>Итоги</h3>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "minmax(0, 2fr) minmax(220px, 1fr)",
                gap: "0.75rem",
                alignItems: "start",
              }}
            >
              <div style={{ display: "grid", gap: "0.75rem" }}>
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))",
                    gap: "0.75rem",
                  }}
                >
                  <FinanceCard
                    label="Общий вес груза, кг"
                    value={formatNumberLocale(Math.max(0, offer.total_cargo_weight_kg ?? 0))}
                  />
                  <FinanceCard label="НДС (22%)" value={formatMoney(offer.finance.vat_amount)} />
                </div>

                <div style={{ border: "1px solid #e4e7ec", borderRadius: 12, padding: "0.9rem", background: "#f8fafc" }}>
                  <FieldWrapper label="Стоимость рейса" error={logisticsError}>
                    <div style={{ position: "relative" }}>
                      <input
                        value={logisticsDraft}
                        onChange={(event) => setLogisticsDraft(event.target.value)}
                        inputMode="decimal"
                        placeholder="Стоимость одного рейса"
                        disabled={financePending}
                        style={{
                          width: "100%",
                          border: "1px solid #d0d5dd",
                          borderRadius: 12,
                          padding: "0.8rem 3.5rem 0.8rem 0.9rem",
                          background: "#ffffff",
                          boxSizing: "border-box",
                        }}
                      />
                      <Button
                        type="button"
                        variant="secondary"
                        onClick={handleLogisticsOk}
                        disabled={financePending}
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

                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: "0.75rem",
                    border: "1px solid #e4e7ec",
                    borderRadius: 12,
                    padding: "0.85rem 1rem",
                    background: "#fafafa",
                  }}
                >
                  <span style={{ color: "#475467", fontWeight: 500 }}>
                    Услуга по доставке грузов
                    {clientTrips > 0 ? ` (${tripsRussianLabel(clientTrips)})` : ""}
                  </span>
                  <strong style={{ fontVariantNumeric: "tabular-nums" }}>{formatMoney(offer.delivery_service_total_rub)}</strong>
                </div>
              </div>

              <div style={{ display: "grid", gap: "0.75rem", alignSelf: "stretch" }}>
                <div style={{ border: "1px solid #e4e7ec", borderRadius: 12, padding: "0.9rem", background: "#f8fafc" }}>
                  <FieldWrapper label="Целевая сумма (₽)" error={targetSumError}>
                    <div style={{ position: "relative" }}>
                      <input
                        value={targetSumDraft}
                        onChange={(event) => handleTargetSumChange(event.target.value)}
                        inputMode="decimal"
                        placeholder="Например, 2 000 000"
                        disabled={financePending || baseProducts <= 0}
                        style={{
                          width: "100%",
                          border: "1px solid #d0d5dd",
                          borderRadius: 12,
                          padding: "0.8rem 3.5rem 0.8rem 0.9rem",
                          background: "#ffffff",
                          boxSizing: "border-box",
                        }}
                      />
                      <Button
                        type="button"
                        variant="secondary"
                        onClick={handleTargetSumOk}
                        disabled={financePending || baseProducts <= 0}
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
                <div style={{ border: "1px solid #e4e7ec", borderRadius: 12, padding: "0.9rem", background: "#f8fafc" }}>
                  <FieldWrapper label="Скидка (%)" error={discountError}>
                    <div style={{ position: "relative" }}>
                      <input
                        value={discountDraft}
                        onChange={(event) => {
                          setDiscountDraft(event.target.value);
                          const discount = parseNumberField(event.target.value);
                          if (discount === null) {
                            return;
                          }
                          const target = targetSumFromDiscountPercent({
                            discountPercent: discount,
                            baseProductsTotalWithVat: baseProducts,
                            deliveryTotal,
                          });
                          if (target !== null) {
                            setTargetSumDraft(String(target).replace(".", ","));
                            setTargetSumError(null);
                          }
                        }}
                        inputMode="decimal"
                        placeholder="Например, 5"
                        disabled={financePending}
                        style={{
                          width: "100%",
                          border: "1px solid #d0d5dd",
                          borderRadius: 12,
                          padding: "0.8rem 3.5rem 0.8rem 0.9rem",
                          background: "#ffffff",
                          boxSizing: "border-box",
                        }}
                      />
                      <Button
                        type="button"
                        variant="secondary"
                        onClick={handleDiscountOk}
                        disabled={financePending}
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
                <FinanceCard label="Итого с НДС" value={formatMoney(offer.finance.total_amount)} accent />
              </div>
            </div>

            {(discountMutation.isError || logisticsMutation.isError) && (
              <div style={{ marginTop: "0.75rem" }}>
                <Alert tone="error">{getErrorMessage(discountMutation.error ?? logisticsMutation.error)}</Alert>
              </div>
            )}
          </section>

          <section>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "0.5rem" }}>
              <h3 style={{ margin: 0 }}>Состав заказа ({orderItemsCount})</h3>
              {orderItemsCount > PLATES_PREVIEW && (
                <Button variant="ghost" onClick={() => setShowAllPlates((prev) => !prev)}>
                  {showAllPlates ? "Свернуть" : "Показать все"}
                </Button>
              )}
            </div>
            {orderItemsCount === 0 ? (
              <div style={{ color: "#667085" }}>
                {isStepOffer ? "Список ступеней пуст." : isMarchOffer ? "Список маршей пуст." : isPileOffer ? "Список свай пуст." : "Список плит пуст."}
              </div>
            ) : isStepOffer ? (
              <div style={{ overflowX: "auto" }}>
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "0.95rem" }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                      <th style={{ padding: "0.5rem 0.75rem" }}>№</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Марка</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Кол-во</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Цена</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(itemsToShow as ArchiveStepItem[]).map((step, index) => (
                      <tr key={`${step.mark}-${index}`} style={{ borderTop: "1px solid #e4e7ec" }}>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{step.position_number ?? index + 1}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{step.mark || "—"}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{step.qty} шт</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>
                          {formatLinePrice(step.discounted_price, step.unit_price)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : isMarchOffer ? (
              <div style={{ overflowX: "auto" }}>
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "0.95rem" }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                      <th style={{ padding: "0.5rem 0.75rem" }}>№</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Марка</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Класс</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Кол-во</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Цена</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(itemsToShow as ArchiveMarchItem[]).map((march, index) => (
                      <tr key={`${march.mark}-${index}`} style={{ borderTop: "1px solid #e4e7ec" }}>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{march.position_number ?? index + 1}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{march.mark || "—"}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{march.concrete_grade || "—"}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{march.qty} шт</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>
                          {formatLinePrice(march.discounted_price, march.unit_price)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : isPileOffer || isBridgePileOffer || isFbsOffer ? (
              <div style={{ overflowX: "auto" }}>
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "0.95rem" }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                      <th style={{ padding: "0.5rem 0.75rem" }}>№</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Марка</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Класс</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Кол-во</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Цена</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(itemsToShow as ArchivePileItem[]).map((pile, index) => (
                      <tr key={`${pile.mark}-${index}`} style={{ borderTop: "1px solid #e4e7ec" }}>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{pile.position_number ?? index + 1}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{pile.mark || "—"}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{pile.concrete_grade || "—"}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{pile.qty} шт</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>
                          {formatLinePrice(pile.discounted_price, pile.unit_price)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div style={{ overflowX: "auto" }}>
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "0.95rem" }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                      <th style={{ padding: "0.5rem 0.75rem" }}>№</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Наименование</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Кол-во</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Цена</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(itemsToShow as typeof offer.plates).map((plate, index) => (
                      <tr key={`${plate.plate_name}-${index}`} style={{ borderTop: "1px solid #e4e7ec" }}>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{plate.position_number ?? index + 1}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{plate.plate_name || "—"}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{plate.qty} шт</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>
                          {formatLinePrice(plate.discounted_price, plate.unit_price)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>

          <section style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
            <Button onClick={() => downloadFile(archiveApi.buildDocumentUrl(offer.kp_id, "pdf"))}>📄 PDF</Button>
            <Button onClick={() => downloadFile(archiveApi.buildDocumentUrl(offer.kp_id, "xlsx"))}>📊 XLSX</Button>
            {!isSimpleProductOffer && (
              <Button
                variant="secondary"
                disabled={schemaMutation.isPending}
                onClick={() => schemaMutation.mutate(offer.kp_id)}
              >
                {schemaMutation.isPending ? "Формируем…" : "📐 Схема"}
              </Button>
            )}
            {offer.status === "в работе" && (
              <Button
                variant="secondary"
                disabled={resumePending}
                onClick={() => {
                  void handleResumeAppend();
                }}
              >
                {resumePending ? "Открываем…" : "Добавить другое наименование"}
              </Button>
            )}
            {canShowDeliverySchedule && (
              <Button
                variant="secondary"
                title={
                  canEditDeliverySchedule
                    ? "Редактирование графика поставки"
                    : "Просмотр графика поставки"
                }
                onClick={() => setScheduleOpen(true)}
              >
                График поставки
              </Button>
            )}
            {offer.status === "в архиве" && (
              isSimpleProductOffer ? (
                <Button variant="secondary" disabled title="скоро">
                  🏭 В производство
                </Button>
              ) : (
                <Button variant="secondary" onClick={() => setMoveOpen(true)}>
                  🏭 В производство
                </Button>
              )
            )}
            <Button variant="danger" onClick={() => setDeleteOpen(true)}>
              Удалить КП
            </Button>
          </section>

          {schemaMutation.isError && (
            <Alert tone="error">{getErrorMessage(schemaMutation.error)}</Alert>
          )}
          {resumeError && <Alert tone="error">{resumeError}</Alert>}
        </div>
      )}

      {offer && (
        <>
          <DeleteConfirmDialog
            open={deleteOpen}
            onClose={() => setDeleteOpen(false)}
            onDeleted={onClose}
            kpId={offer.kp_id}
            customerName={offer.customer_name}
          />
          <MoveToProductionDialog
            open={moveOpen}
            onClose={() => setMoveOpen(false)}
            kpId={offer.kp_id}
            initialExecutionTerms={offer.execution_terms}
          />
          <HighDiscountConfirmDialog
            open={pendingDiscountPercent !== null}
            discountPercent={pendingDiscountPercent ?? 0}
            isPending={discountMutation.isPending}
            onConfirm={() => pendingDiscountPercent !== null && void applyDiscount(pendingDiscountPercent)}
            onCancel={() => {
              setPendingDiscountPercent(null);
              restoreDiscountDrafts();
            }}
          />
          {canShowDeliverySchedule && (
            <DeliveryScheduleDialog
              open={scheduleOpen}
              onClose={() => setScheduleOpen(false)}
              kpId={offer.kp_id}
              plates={schedulePlates}
              readOnly={!canEditDeliverySchedule}
            />
          )}
        </>
      )}
    </Modal>
  );
};

const FinanceCard = ({ label, value, accent }: { label: string; value: string; accent?: boolean }) => (
  <div
    style={{
      padding: "0.75rem 1rem",
      borderRadius: 14,
      border: "1px solid #e4e7ec",
      background: accent ? "#eef4ff" : "#ffffff",
    }}
  >
    <div style={{ fontSize: "0.85rem", color: "#667085" }}>{label}</div>
    <div style={{ fontWeight: 700, color: accent ? "#1d4ed8" : "#101828", marginTop: "0.25rem" }}>{value}</div>
  </div>
);

const parseNumberField = (raw: string): number | null => {
  const normalized = raw.trim().replace(/\s+/g, "").replace(",", ".");
  if (!normalized.length) {
    return null;
  }
  const n = Number(normalized);
  return Number.isFinite(n) ? n : null;
};

const formatNumberLocale = (value: number): string =>
  value.toLocaleString("ru-RU", { maximumFractionDigits: 2 });

/** Подпись «N рейс/рейса/рейсов» для подсказки к расчётной строке доставки. */
const tripsRussianLabel = (n: number): string => {
  const k = Math.abs(Math.trunc(n));
  const mod100 = k % 100;
  const mod10 = k % 10;
  if (mod100 >= 11 && mod100 <= 14) {
    return `${k} рейсов`;
  }
  if (mod10 === 1) {
    return `${k} рейс`;
  }
  if (mod10 >= 2 && mod10 <= 4) {
    return `${k} рейса`;
  }
  return `${k} рейсов`;
};
