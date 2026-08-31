import { useEffect, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Drawer } from "@/shared/ui/Drawer";
import { Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { Spinner } from "@/shared/ui/Spinner";
import { getErrorMessage } from "@/shared/lib/apiError";
import { ShipmentFieldsSection } from "@/features/logistics/components/ShipmentFieldsSection";
import { ShipmentItemsSection } from "@/features/logistics/components/ShipmentItemsSection";
import {
  useCancelShipmentMutation,
  useCompleteShipmentMutation,
  useShipmentQuery,
  useShipmentSheetMutation,
  useUpdateShipmentMutation,
} from "@/features/logistics/hooks/useLogisticsQueries";
import {
  deliveryTypeLabel,
  formatDate,
  shipmentStatusLabel,
} from "@/features/logistics/lib/logisticsFormat";
import type {
  DeliveryType,
  ShipmentDetails,
  UpdateShipmentPayload,
} from "@/features/logistics/types/logistics";

type Props = {
  shipmentId: number | null;
  onClose: () => void;
};

type OrderDraft = {
  kp_id: number;
  ya_order_no: string;
  customer_name: string | null;
};

type HeaderDraft = {
  shipment_date: string;
  delivery_type: DeliveryType;
  attention: boolean;
  attention_comment: string;
  orders: OrderDraft[];
};

const headerFromShipment = (shipment: ShipmentDetails): HeaderDraft => ({
  shipment_date: shipment.shipment_date.slice(0, 10),
  delivery_type: shipment.delivery_type,
  attention: Boolean(shipment.attention),
  attention_comment: shipment.attention_comment ?? "",
  orders: shipment.orders.map((order) => ({
    kp_id: order.kp_id,
    ya_order_no: order.ya_order_no ?? "",
    customer_name: order.customer_name,
  })),
});

const headerVersion = (shipment: ShipmentDetails): string =>
  JSON.stringify([
    shipment.shipment_date,
    shipment.delivery_type,
    shipment.attention,
    shipment.attention_comment,
    shipment.orders.map((o) => [o.kp_id, o.ya_order_no]),
  ]);

const sectionCardStyle: React.CSSProperties = {
  border: "1px solid #eaecf0",
  borderRadius: 14,
  background: "#ffffff",
  padding: "1rem",
};

export const ShipmentDrawer = ({ shipmentId, onClose }: Props) => {
  const query = useShipmentQuery(shipmentId);
  const shipment = query.data ?? null;

  const [header, setHeader] = useState<HeaderDraft | null>(null);
  const [headerDirty, setHeaderDirty] = useState(false);
  const [newKpId, setNewKpId] = useState("");
  const [headerError, setHeaderError] = useState<string | null>(null);
  const [actionError, setActionError] = useState<string | null>(null);
  const [confirmCancelOpen, setConfirmCancelOpen] = useState(false);
  const [confirmCompleteOpen, setConfirmCompleteOpen] = useState(false);

  const updateMutation = useUpdateShipmentMutation(shipmentId ?? -1);
  const completeMutation = useCompleteShipmentMutation(shipmentId ?? -1);
  const cancelMutation = useCancelShipmentMutation(shipmentId ?? -1);
  const sheetMutation = useShipmentSheetMutation(shipmentId ?? -1);

  const version = shipment ? headerVersion(shipment) : "";
  useEffect(() => {
    if (shipment) {
      setHeader(headerFromShipment(shipment));
      setHeaderDirty(false);
      setNewKpId("");
    } else {
      setHeader(null);
      setHeaderDirty(false);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [version, shipment?.id]);

  useEffect(() => {
    if (shipmentId === null) {
      setActionError(null);
      setHeaderError(null);
      setConfirmCancelOpen(false);
      setConfirmCompleteOpen(false);
    }
  }, [shipmentId]);

  const patchHeader = (partial: Partial<HeaderDraft>) => {
    setHeader((prev) => (prev ? { ...prev, ...partial } : prev));
    setHeaderDirty(true);
  };

  const patchOrder = (kpId: number, yaOrderNo: string) => {
    setHeader((prev) =>
      prev
        ? {
            ...prev,
            orders: prev.orders.map((order) =>
              order.kp_id === kpId ? { ...order, ya_order_no: yaOrderNo } : order,
            ),
          }
        : prev,
    );
    setHeaderDirty(true);
  };

  const removeOrder = (kpId: number) => {
    setHeader((prev) =>
      prev && prev.orders.length > 1
        ? { ...prev, orders: prev.orders.filter((order) => order.kp_id !== kpId) }
        : prev,
    );
    setHeaderDirty(true);
  };

  const addOrder = () => {
    const kpId = Number(newKpId.trim());
    if (!Number.isInteger(kpId) || kpId <= 0) {
      setHeaderError("Укажите номер КП (целое число больше 0).");
      return;
    }
    if (header?.orders.some((order) => order.kp_id === kpId)) {
      setHeaderError(`КП №${kpId} уже в рейсе.`);
      return;
    }
    setHeaderError(null);
    setHeader((prev) =>
      prev
        ? { ...prev, orders: [...prev.orders, { kp_id: kpId, ya_order_no: "", customer_name: null }] }
        : prev,
    );
    setHeaderDirty(true);
    setNewKpId("");
  };

  const saveHeader = async () => {
    if (!header) return;
    setHeaderError(null);
    if (!header.shipment_date) {
      setHeaderError("Укажите дату рейса.");
      return;
    }
    if (header.orders.length === 0) {
      setHeaderError("В рейсе должен быть хотя бы один заказ (КП).");
      return;
    }
    const payload: UpdateShipmentPayload = {
      shipment_date: header.shipment_date,
      delivery_type: header.delivery_type,
      attention: header.attention,
      attention_comment: header.attention ? header.attention_comment.trim() || null : null,
      orders: header.orders.map((order) => ({
        kp_id: order.kp_id,
        ya_order_no: order.ya_order_no.trim() || null,
      })),
    };
    try {
      await updateMutation.mutateAsync(payload);
      setHeaderDirty(false);
    } catch (err) {
      setHeaderError(getErrorMessage(err));
    }
  };

  const runComplete = async () => {
    setActionError(null);
    try {
      await completeMutation.mutateAsync();
      setConfirmCompleteOpen(false);
    } catch (err) {
      setConfirmCompleteOpen(false);
      setActionError(getErrorMessage(err));
    }
  };

  const runCancel = async () => {
    setActionError(null);
    try {
      await cancelMutation.mutateAsync();
      setConfirmCancelOpen(false);
      onClose();
    } catch (err) {
      setConfirmCancelOpen(false);
      setActionError(getErrorMessage(err));
    }
  };

  const downloadSheet = async () => {
    setActionError(null);
    try {
      await sheetMutation.mutateAsync();
    } catch (err) {
      setActionError(getErrorMessage(err));
    }
  };

  const inWork = shipment?.status === "in_work";
  const missingYaOrders =
    shipment?.orders.filter((order) => !order.ya_order_no?.trim()) ?? [];
  const busy =
    updateMutation.isPending || completeMutation.isPending || cancelMutation.isPending;

  return (
    <Drawer
      open={shipmentId !== null}
      onClose={onClose}
      width={940}
      title={shipment ? `Рейс №${shipment.id} от ${formatDate(shipment.shipment_date)}` : "Рейс"}
    >
      {query.isLoading && <Spinner />}
      {query.isError && <Alert tone="error">{getErrorMessage(query.error)}</Alert>}

      {shipment && header && (
        <div style={{ display: "grid", gap: "1rem" }}>
          <section style={{ ...sectionCardStyle, display: "grid", gap: "0.75rem" }}>
            <div
              style={{
                display: "flex",
                gap: "0.75rem",
                flexWrap: "wrap",
                alignItems: "flex-end",
              }}
            >
              <label style={{ display: "grid", gap: 4, fontSize: "0.85rem", color: "#475467" }}>
                Дата рейса
                <div style={{ width: 160 }}>
                  <Input
                    type="date"
                    value={header.shipment_date}
                    onChange={(e) => patchHeader({ shipment_date: e.target.value })}
                    disabled={!inWork}
                  />
                </div>
              </label>
              <label style={{ display: "grid", gap: 4, fontSize: "0.85rem", color: "#475467" }}>
                Тип выдачи
                <select
                  value={header.delivery_type}
                  onChange={(e) =>
                    patchHeader({ delivery_type: e.target.value as DeliveryType })
                  }
                  disabled={!inWork}
                  style={{
                    border: "1px solid #d0d5dd",
                    borderRadius: 12,
                    padding: "0.8rem 0.9rem",
                    background: "#ffffff",
                  }}
                >
                  <option value="delivery">{deliveryTypeLabel("delivery")}</option>
                  <option value="pickup">{deliveryTypeLabel("pickup")}</option>
                </select>
              </label>
              <span
                style={{
                  fontSize: "0.8rem",
                  fontWeight: 700,
                  padding: "0.3rem 0.7rem",
                  borderRadius: 999,
                  background: shipment.status === "done" ? "#ecfdf3" : "#eef4ff",
                  color: shipment.status === "done" ? "#027a48" : "#1d4ed8",
                }}
              >
                {shipmentStatusLabel(shipment.status)}
              </span>
              <label
                style={{
                  display: "inline-flex",
                  alignItems: "center",
                  gap: 6,
                  fontSize: "0.9rem",
                  color: "#b54708",
                  fontWeight: 600,
                }}
              >
                <input
                  type="checkbox"
                  checked={header.attention}
                  onChange={(e) => patchHeader({ attention: e.target.checked })}
                  disabled={!inWork}
                />
                Внимание
              </label>
            </div>

            {header.attention && (
              <Input
                type="text"
                value={header.attention_comment}
                placeholder="Комментарий («Работа крана!», «БЕЗ ЦЕН!»)"
                onChange={(e) => patchHeader({ attention_comment: e.target.value })}
                disabled={!inWork}
              />
            )}

            <div style={{ display: "grid", gap: "0.45rem" }}>
              <span style={{ fontWeight: 600, fontSize: "0.9rem" }}>Заказы в рейсе</span>
              {header.orders.map((order) => (
                <div
                  key={order.kp_id}
                  style={{
                    display: "flex",
                    gap: "0.6rem",
                    alignItems: "center",
                    flexWrap: "wrap",
                    border: "1px solid #f2f4f7",
                    borderRadius: 10,
                    padding: "0.45rem 0.6rem",
                  }}
                >
                  <span style={{ fontWeight: 600 }}>
                    КП №{order.kp_id}
                    {order.customer_name ? ` · ${order.customer_name}` : ""}
                  </span>
                  <label
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 6,
                      fontSize: "0.85rem",
                      color: "#475467",
                      flex: "1 1 220px",
                    }}
                  >
                    № заказа (ЯР)
                    <Input
                      type="text"
                      value={order.ya_order_no}
                      placeholder="ЯР-0000000"
                      onChange={(e) => patchOrder(order.kp_id, e.target.value)}
                      disabled={!inWork}
                    />
                  </label>
                  {inWork && header.orders.length > 1 && (
                    <Button
                      variant="ghost"
                      aria-label={`Убрать КП №${order.kp_id}`}
                      onClick={() => removeOrder(order.kp_id)}
                      disabled={busy}
                    >
                      ×
                    </Button>
                  )}
                </div>
              ))}
              {inWork && (
                <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
                  <div style={{ width: 180 }}>
                    <Input
                      type="number"
                      min={1}
                      placeholder="Номер КП"
                      value={newKpId}
                      onChange={(e) => setNewKpId(e.target.value)}
                    />
                  </div>
                  <Button variant="secondary" onClick={addOrder} disabled={busy}>
                    Добавить КП
                  </Button>
                </div>
              )}
            </div>

            {headerError && <Alert tone="error">{headerError}</Alert>}
            {inWork && headerDirty && (
              <div>
                <Button onClick={saveHeader} disabled={busy}>
                  {updateMutation.isPending ? "Сохранение..." : "Сохранить шапку"}
                </Button>
              </div>
            )}
          </section>

          <div style={sectionCardStyle}>
            <ShipmentItemsSection shipment={shipment} readOnly={!inWork} />
          </div>

          <div style={sectionCardStyle}>
            <ShipmentFieldsSection shipment={shipment} readOnly={!inWork} />
          </div>

          {actionError && <Alert tone="error">{actionError}</Alert>}

          <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
            <Button variant="secondary" onClick={downloadSheet} disabled={sheetMutation.isPending}>
              {sheetMutation.isPending ? "Формирование..." : "Лист отгрузки (XLSX)"}
            </Button>
            {inWork && (
              <>
                <Button
                  variant="danger"
                  onClick={() => setConfirmCancelOpen(true)}
                  disabled={busy}
                >
                  Отменить рейс
                </Button>
                <Button
                  onClick={() => {
                    setActionError(null);
                    if (headerDirty) {
                      setActionError("Сначала сохраните шапку — есть несохранённые изменения.");
                      return;
                    }
                    if (missingYaOrders.length > 0) {
                      setActionError(
                        `Заполните № заказа (ЯР) для: ${missingYaOrders
                          .map((order) => `КП №${order.kp_id}`)
                          .join(", ")}.`,
                      );
                      return;
                    }
                    setConfirmCompleteOpen(true);
                  }}
                  disabled={busy}
                >
                  {completeMutation.isPending ? "Закрытие..." : "Выезд"}
                </Button>
              </>
            )}
          </div>
        </div>
      )}

      <Modal
        open={confirmCancelOpen}
        onClose={() => setConfirmCancelOpen(false)}
        title="Отменить рейс?"
        maxWidth={440}
      >
        <p style={{ marginTop: 0, color: "#475467" }}>
          Состав рейса будет удалён, зарезервированные плиты вернутся в свободные на СГП.
        </p>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button variant="ghost" onClick={() => setConfirmCancelOpen(false)} disabled={cancelMutation.isPending}>
            Назад
          </Button>
          <Button variant="danger" onClick={runCancel} disabled={cancelMutation.isPending}>
            {cancelMutation.isPending ? "Отмена..." : "Да, отменить рейс"}
          </Button>
        </div>
      </Modal>

      <Modal
        open={confirmCompleteOpen}
        onClose={() => setConfirmCompleteOpen(false)}
        title="Подтвердить выезд?"
        maxWidth={440}
      >
        <p style={{ marginTop: 0, color: "#475467" }}>
          Плиты состава будут списаны со склада СГП, рейс перейдёт в «Обработано». Действие
          необратимо.
        </p>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button variant="ghost" onClick={() => setConfirmCompleteOpen(false)} disabled={completeMutation.isPending}>
            Назад
          </Button>
          <Button onClick={runComplete} disabled={completeMutation.isPending}>
            {completeMutation.isPending ? "Закрытие..." : "Да, выезд"}
          </Button>
        </div>
      </Modal>
    </Drawer>
  );
};
