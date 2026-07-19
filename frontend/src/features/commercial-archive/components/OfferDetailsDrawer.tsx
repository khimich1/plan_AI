import { useEffect, useState } from "react";
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
import { downloadFile } from "@/shared/lib/downloadFile";
import { getErrorMessage } from "@/shared/lib/apiError";
import { DeleteConfirmDialog } from "./DeleteConfirmDialog";
import { MoveToProductionDialog } from "./MoveToProductionDialog";

type Props = {
  open: boolean;
  kpId: number | null;
  onClose: () => void;
};

const PLATES_PREVIEW = 10;

export const OfferDetailsDrawer = ({ open, kpId, onClose }: Props) => {
  const query = useArchiveOfferQuery(open ? kpId : null);
  const [showAllPlates, setShowAllPlates] = useState(false);
  const [moveOpen, setMoveOpen] = useState(false);
  const [discountDraft, setDiscountDraft] = useState("");
  const [logisticsDraft, setLogisticsDraft] = useState("");
  const [discountError, setDiscountError] = useState<string | null>(null);
  const [logisticsError, setLogisticsError] = useState<string | null>(null);
  const [deleteOpen, setDeleteOpen] = useState(false);

  const discountMutation = useUpdateDiscountMutation();
  const logisticsMutation = useUpdateLogisticsCostMutation();
  const schemaMutation = useArchiveDocumentMutation("schema");
  const financePending = discountMutation.isPending || logisticsMutation.isPending;

  const offer = query.data;
  const platesToShow = offer
    ? showAllPlates
      ? offer.plates
      : offer.plates.slice(0, PLATES_PREVIEW)
    : [];

  useEffect(() => {
    if (offer) {
      setDiscountDraft(String(offer.finance.discount_percent ?? 0));
      setLogisticsDraft(String(offer.logistics_cost ?? 0).replace(".", ","));
      setDiscountError(null);
      setLogisticsError(null);
    }
  }, [offer?.kp_id, offer?.finance.discount_percent, offer?.logistics_cost]);

  const handleDiscountOk = async () => {
    if (!offer) {
      return;
    }
    const parsed = parseNumberField(discountDraft);
    if (parsed === null || parsed < 0 || parsed > 100) {
      setDiscountError("Скидка должна быть числом от 0 до 100.");
      return;
    }
    setDiscountError(null);
    await discountMutation.mutateAsync({ kpId: offer.kp_id, discount: parsed });
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
    <Modal open={open} onClose={onClose} title={offer ? `КП №${offer.kp_id}` : "Карточка КП"} maxWidth={720}>
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
              {offer.completion_percentage !== null && (
                <div>
                  <div style={{ color: "#667085", fontSize: "0.85rem" }}>Готовность</div>
                  <div style={{ fontWeight: 600 }}>{offer.completion_percentage.toFixed(1)}%</div>
                </div>
              )}
            </div>
          </section>

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
                  <FieldWrapper label="Скидка (%)" error={discountError}>
                    <div style={{ position: "relative" }}>
                      <input
                        value={discountDraft}
                        onChange={(event) => setDiscountDraft(event.target.value)}
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
              <h3 style={{ margin: 0 }}>Состав заказа ({offer.plates.length})</h3>
              {offer.plates.length > PLATES_PREVIEW && (
                <Button variant="ghost" onClick={() => setShowAllPlates((prev) => !prev)}>
                  {showAllPlates ? "Свернуть" : "Показать все"}
                </Button>
              )}
            </div>
            {offer.plates.length === 0 ? (
              <div style={{ color: "#667085" }}>Список плит пуст.</div>
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
                    {platesToShow.map((plate, index) => (
                      <tr key={`${plate.plate_name}-${index}`} style={{ borderTop: "1px solid #e4e7ec" }}>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{plate.position_number ?? index + 1}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{plate.plate_name || "—"}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{plate.qty} шт</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>
                          {plate.discounted_price !== null
                            ? formatMoney(plate.discounted_price)
                            : plate.unit_price !== null
                              ? formatMoney(plate.unit_price)
                              : "—"}
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
            <Button
              variant="secondary"
              disabled={schemaMutation.isPending}
              onClick={() => schemaMutation.mutate(offer.kp_id)}
            >
              {schemaMutation.isPending ? "Формируем…" : "📐 Схема"}
            </Button>
            {offer.status === "в архиве" && (
              <Button variant="secondary" onClick={() => setMoveOpen(true)}>
                🏭 В производство
              </Button>
            )}
            <Button variant="danger" onClick={() => setDeleteOpen(true)}>
              Удалить КП
            </Button>
          </section>

          {schemaMutation.isError && (
            <Alert tone="error">{getErrorMessage(schemaMutation.error)}</Alert>
          )}
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
