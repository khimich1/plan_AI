import { useMemo, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import { getErrorMessage } from "@/shared/lib/apiError";
import { CarrierAutocomplete } from "@/features/logistics/components/CarrierAutocomplete";
import { CreateShipmentDialog } from "@/features/logistics/components/CreateShipmentDialog";
import { ShipmentDrawer } from "@/features/logistics/components/ShipmentDrawer";
import { useShipmentsQuery } from "@/features/logistics/hooks/useLogisticsQueries";
import {
  deliveryTypeLabel,
  deliveryTypeShort,
  formatCost,
  formatDate,
  formatWeightKg,
  isFlagOn,
  shipmentStatusLabel,
} from "@/features/logistics/lib/logisticsFormat";
import type {
  DeliveryType,
  ShipmentFilters,
  ShipmentRegistryRow,
} from "@/features/logistics/types/logistics";

type RegistryFilterState = {
  dateFrom: string;
  dateTo: string;
  orderQuery: string;
  carrier: { id: number; name: string } | null;
  deliveryType: "" | DeliveryType;
  noUpd: boolean;
  attentionOnly: boolean;
};

const EMPTY_FILTERS: RegistryFilterState = {
  dateFrom: "",
  dateTo: "",
  orderQuery: "",
  carrier: null,
  deliveryType: "",
  noUpd: false,
  attentionOnly: false,
};

export const toServerFilters = (state: RegistryFilterState): ShipmentFilters => {
  const orderRaw = state.orderQuery.trim();
  const orderNumeric = Number(orderRaw);
  return {
    date_from: state.dateFrom || undefined,
    date_to: state.dateTo || undefined,
    kp_id:
      orderRaw && Number.isInteger(orderNumeric) && orderNumeric > 0 ? orderNumeric : undefined,
    carrier_id: state.carrier?.id,
    delivery_type: state.deliveryType || undefined,
    no_upd: state.noUpd || undefined,
    attention: state.attentionOnly || undefined,
  };
};

/** Текстовый остаток «заказа» (ЯР/заказчик) фильтруем на клиенте — в контракте только kp_id. */
export const applyClientOrderFilter = (
  rows: ShipmentRegistryRow[],
  orderQuery: string,
): ShipmentRegistryRow[] => {
  const raw = orderQuery.trim().toLowerCase();
  if (!raw || (Number.isInteger(Number(raw)) && Number(raw) > 0)) {
    return rows;
  }
  return rows.filter((row) =>
    row.orders.some(
      (order) =>
        (order.ya_order_no ?? "").toLowerCase().includes(raw) ||
        (order.customer_name ?? "").toLowerCase().includes(raw),
    ),
  );
};

const ordersStackLabel = (row: ShipmentRegistryRow): string[] =>
  row.orders.map((order) => order.ya_order_no?.trim() || `КП №${order.kp_id}`);

const customersLabel = (row: ShipmentRegistryRow): string => {
  const names = new Set<string>();
  for (const order of row.orders) {
    if (order.customer_name?.trim()) {
      names.add(order.customer_name.trim());
    }
  }
  return names.size > 0 ? [...names].join(", ") : "—";
};

const carrierLabel = (row: ShipmentRegistryRow): string => {
  if (row.delivery_type === "pickup") {
    return row.proxy_no?.trim() ? `дов. ${row.proxy_no.trim()}` : "—";
  }
  return row.carrier_name?.trim() || "—";
};

const statusBadgeStyle = (status: ShipmentRegistryRow["status"]): React.CSSProperties => ({
  display: "inline-block",
  fontSize: "0.78rem",
  fontWeight: 700,
  padding: "0.15rem 0.55rem",
  borderRadius: 999,
  background: status === "done" ? "#ecfdf3" : "#eef4ff",
  color: status === "done" ? "#027a48" : "#1d4ed8",
});

export const LogisticsRegistryView = () => {
  const [filterState, setFilterState] = useState<RegistryFilterState>(EMPTY_FILTERS);
  const serverFilters = useMemo(() => toServerFilters(filterState), [filterState]);
  const query = useShipmentsQuery(serverFilters);
  const rows = useMemo(
    () => applyClientOrderFilter(query.data ?? [], filterState.orderQuery),
    [query.data, filterState.orderQuery],
  );

  const [selectedId, setSelectedId] = useState<number | null>(null);
  const [createOpen, setCreateOpen] = useState(false);
  const [reuseSource, setReuseSource] = useState<{
    id: number;
    delivery_type: DeliveryType;
  } | null>(null);

  const patch = (partial: Partial<RegistryFilterState>) =>
    setFilterState((prev) => ({ ...prev, ...partial }));

  const openCreate = () => {
    setReuseSource(null);
    setCreateOpen(true);
  };

  const openReuse = (row: ShipmentRegistryRow) => {
    setReuseSource({ id: row.id, delivery_type: row.delivery_type });
    setCreateOpen(true);
  };

  const closeCreate = () => {
    setCreateOpen(false);
    setReuseSource(null);
  };

  const hasActiveFilters =
    JSON.stringify(toServerFilters(filterState)) !== JSON.stringify(toServerFilters(EMPTY_FILTERS)) ||
    (filterState.orderQuery.trim() !== "" && toServerFilters(filterState).kp_id == null);

  return (
    <section style={{ display: "grid", gap: "1rem" }}>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          gap: "0.75rem",
          flexWrap: "wrap",
        }}
      >
        <h2 style={{ margin: 0, fontSize: "1.25rem" }}>Реестр рейсов</h2>
        <Button onClick={openCreate}>Новый рейс</Button>
      </div>

      <div
        style={{
          display: "flex",
          gap: "0.6rem",
          flexWrap: "wrap",
          alignItems: "flex-end",
          border: "1px solid #eaecf0",
          borderRadius: 14,
          background: "#ffffff",
          padding: "0.9rem 1rem",
        }}
      >
        <label style={{ display: "grid", gap: 4, fontSize: "0.85rem", color: "#475467" }}>
          Дата с
          <div style={{ width: 150 }}>
            <Input
              type="date"
              value={filterState.dateFrom}
              onChange={(e) => patch({ dateFrom: e.target.value })}
            />
          </div>
        </label>
        <label style={{ display: "grid", gap: 4, fontSize: "0.85rem", color: "#475467" }}>
          Дата по
          <div style={{ width: 150 }}>
            <Input
              type="date"
              value={filterState.dateTo}
              onChange={(e) => patch({ dateTo: e.target.value })}
            />
          </div>
        </label>
        <label style={{ display: "grid", gap: 4, fontSize: "0.85rem", color: "#475467", flex: "1 1 180px" }}>
          Заказ (№ КП / ЯР / заказчик)
          <Input
            type="text"
            placeholder="например 154 или ЯР-0001"
            value={filterState.orderQuery}
            onChange={(e) => patch({ orderQuery: e.target.value })}
          />
        </label>
        <label style={{ display: "grid", gap: 4, fontSize: "0.85rem", color: "#475467", flex: "1 1 200px" }}>
          Перевозчик
          <CarrierAutocomplete
            selected={filterState.carrier}
            onSelect={(carrier) => patch({ carrier })}
            placeholder="Все перевозчики"
          />
        </label>
        <label style={{ display: "grid", gap: 4, fontSize: "0.85rem", color: "#475467" }}>
          Тип
          <select
            aria-label="Тип"
            value={filterState.deliveryType}
            onChange={(e) => patch({ deliveryType: e.target.value as "" | DeliveryType })}
            style={{
              border: "1px solid #d0d5dd",
              borderRadius: 12,
              padding: "0.8rem 0.9rem",
              background: "#ffffff",
            }}
          >
            <option value="">Все</option>
            <option value="delivery">{deliveryTypeLabel("delivery")}</option>
            <option value="pickup">{deliveryTypeLabel("pickup")}</option>
          </select>
        </label>
        <label
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            fontSize: "0.9rem",
            color: "#344054",
            paddingBottom: 10,
          }}
        >
          <input
            type="checkbox"
            checked={filterState.noUpd}
            onChange={(e) => patch({ noUpd: e.target.checked })}
          />
          Без УПД
        </label>
        <label
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            fontSize: "0.9rem",
            color: "#344054",
            paddingBottom: 10,
          }}
        >
          <input
            type="checkbox"
            checked={filterState.attentionOnly}
            onChange={(e) => patch({ attentionOnly: e.target.checked })}
          />
          Внимание
        </label>
        {hasActiveFilters && (
          <Button variant="ghost" onClick={() => setFilterState(EMPTY_FILTERS)}>
            Сбросить
          </Button>
        )}
      </div>

      {query.isError && <Alert tone="error">{getErrorMessage(query.error)}</Alert>}

      {query.isLoading ? (
        <Spinner />
      ) : rows.length === 0 ? (
        <Alert tone="info">Рейсов по выбранным фильтрам нет.</Alert>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.9rem" }}>
            <thead>
              <tr style={{ textAlign: "left", borderBottom: "1px solid #eaecf0" }}>
                <th style={{ padding: "0.55rem" }}>Дата</th>
                <th style={{ padding: "0.55rem" }}>Заказы</th>
                <th style={{ padding: "0.55rem" }}>Заказчики</th>
                <th style={{ padding: "0.55rem" }} title="Доставка / Самовывоз">
                  Д/С
                </th>
                <th style={{ padding: "0.55rem" }}>Перевозчик / доверенность</th>
                <th style={{ padding: "0.55rem" }}>Водитель</th>
                <th style={{ padding: "0.55rem" }}>ТС</th>
                <th style={{ padding: "0.55rem" }}>Вес</th>
                <th style={{ padding: "0.55rem" }}>УПД</th>
                <th style={{ padding: "0.55rem" }}>Стоимость план</th>
                <th style={{ padding: "0.55rem" }}>Статус</th>
                <th style={{ padding: "0.55rem" }} aria-label="Внимание" />
                <th style={{ padding: "0.55rem" }} aria-label="Действия" />
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr
                  key={row.id}
                  onClick={() => setSelectedId(row.id)}
                  style={{ borderBottom: "1px solid #f2f4f7", cursor: "pointer" }}
                >
                  <td style={{ padding: "0.55rem", whiteSpace: "nowrap" }}>
                    {formatDate(row.shipment_date)}
                  </td>
                  <td style={{ padding: "0.55rem" }}>
                    <div style={{ display: "grid", gap: 2 }}>
                      {ordersStackLabel(row).map((label, index) => (
                        <span key={index} style={{ fontWeight: 600 }}>
                          {label}
                        </span>
                      ))}
                    </div>
                  </td>
                  <td style={{ padding: "0.55rem" }}>{customersLabel(row)}</td>
                  <td style={{ padding: "0.55rem" }}>
                    <span
                      title={deliveryTypeLabel(row.delivery_type)}
                      style={{
                        display: "inline-block",
                        minWidth: 22,
                        textAlign: "center",
                        fontWeight: 700,
                        fontSize: "0.78rem",
                        padding: "0.15rem 0.4rem",
                        borderRadius: 8,
                        background: row.delivery_type === "pickup" ? "#fffaeb" : "#eef4ff",
                        color: row.delivery_type === "pickup" ? "#b54708" : "#1d4ed8",
                      }}
                    >
                      {deliveryTypeShort(row.delivery_type)}
                    </span>
                  </td>
                  <td style={{ padding: "0.55rem" }}>{carrierLabel(row)}</td>
                  <td style={{ padding: "0.55rem" }}>{row.driver_name || "—"}</td>
                  <td style={{ padding: "0.55rem" }}>{row.vehicle_text || "—"}</td>
                  <td style={{ padding: "0.55rem", whiteSpace: "nowrap" }}>
                    {formatWeightKg(row.total_weight_kg)}
                  </td>
                  <td style={{ padding: "0.55rem" }}>{row.upd_no || "—"}</td>
                  <td style={{ padding: "0.55rem", whiteSpace: "nowrap" }}>
                    {formatCost(row.planned_cost)}
                  </td>
                  <td style={{ padding: "0.55rem" }}>
                    <span style={statusBadgeStyle(row.status)}>{shipmentStatusLabel(row.status)}</span>
                  </td>
                  <td style={{ padding: "0.55rem" }}>
                    {isFlagOn(row.attention) && (
                      <span
                        role="img"
                        aria-label="Внимание"
                        title={row.attention_comment?.trim() || "Внимание"}
                        style={{ cursor: "help" }}
                      >
                        ⚠️
                      </span>
                    )}
                  </td>
                  <td style={{ padding: "0.55rem" }}>
                    <Button
                      type="button"
                      variant="ghost"
                      title="Создать на основе"
                      onClick={(event) => {
                        event.stopPropagation();
                        openReuse(row);
                      }}
                    >
                      На основе
                    </Button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <CreateShipmentDialog
        open={createOpen}
        sourceShipmentId={reuseSource?.id}
        initialDeliveryType={reuseSource?.delivery_type}
        onClose={closeCreate}
        onCreated={(id) => {
          closeCreate();
          setSelectedId(id);
        }}
      />
      <ShipmentDrawer shipmentId={selectedId} onClose={() => setSelectedId(null)} />
    </section>
  );
};
