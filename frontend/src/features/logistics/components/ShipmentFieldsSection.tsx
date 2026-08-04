import { useEffect, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { getErrorMessage } from "@/shared/lib/apiError";
import { CarrierAutocomplete } from "@/features/logistics/components/CarrierAutocomplete";
import { useUpdateShipmentMutation } from "@/features/logistics/hooks/useLogisticsQueries";
import type {
  ShipmentDetails,
  UpdateShipmentPayload,
  VehicleClass,
} from "@/features/logistics/types/logistics";

type FieldsDraft = {
  carrier: { id: number; name: string } | null;
  proxy_no: string;
  driver_name: string;
  vehicle_text: string;
  vehicle_class: "" | VehicleClass;
  upd_no: string;
  freight_request_no: string;
  planned_cost: string;
};

const draftFromShipment = (shipment: ShipmentDetails): FieldsDraft => ({
  carrier:
    shipment.carrier_id != null
      ? { id: shipment.carrier_id, name: shipment.carrier_name ?? "" }
      : null,
  proxy_no: shipment.proxy_no ?? "",
  driver_name: shipment.driver_name ?? "",
  vehicle_text: shipment.vehicle_text ?? "",
  vehicle_class: shipment.vehicle_class ?? "",
  upd_no: shipment.upd_no ?? "",
  freight_request_no: shipment.freight_request_no ?? "",
  planned_cost: shipment.planned_cost != null ? String(shipment.planned_cost) : "",
});

const fieldsVersion = (shipment: ShipmentDetails): string =>
  JSON.stringify([
    shipment.carrier_id,
    shipment.proxy_no,
    shipment.driver_name,
    shipment.vehicle_text,
    shipment.vehicle_class,
    shipment.upd_no,
    shipment.freight_request_no,
    shipment.planned_cost,
  ]);

const labelStyle: React.CSSProperties = {
  display: "grid",
  gap: "0.45rem",
  fontSize: "0.9rem",
  color: "#475467",
};

const selectStyle: React.CSSProperties = {
  border: "1px solid #d0d5dd",
  borderRadius: 12,
  padding: "0.8rem 0.9rem",
  background: "#ffffff",
};

export const ShipmentFieldsSection = ({
  shipment,
  readOnly,
}: {
  shipment: ShipmentDetails;
  readOnly: boolean;
}) => {
  const [draft, setDraft] = useState<FieldsDraft>(() => draftFromShipment(shipment));
  const [dirty, setDirty] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const updateMutation = useUpdateShipmentMutation(shipment.id);

  const version = fieldsVersion(shipment);
  useEffect(() => {
    setDraft(draftFromShipment(shipment));
    setDirty(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [version, shipment.id]);

  const patch = (partial: Partial<FieldsDraft>) => {
    setDraft((prev) => ({ ...prev, ...partial }));
    setDirty(true);
  };

  const save = async () => {
    setError(null);
    const plannedCostRaw = draft.planned_cost.trim().replace(",", ".");
    const plannedCost = plannedCostRaw === "" ? null : Number(plannedCostRaw);
    if (plannedCostRaw !== "" && !Number.isFinite(plannedCost)) {
      setError("Стоимость план должна быть числом.");
      return;
    }
    const payload: UpdateShipmentPayload = {
      carrier_id: draft.carrier?.id ?? null,
      proxy_no: draft.proxy_no.trim() || null,
      driver_name: draft.driver_name.trim() || null,
      vehicle_text: draft.vehicle_text.trim() || null,
      vehicle_class: draft.vehicle_class || null,
      upd_no: draft.upd_no.trim() || null,
      freight_request_no: draft.freight_request_no.trim() || null,
      planned_cost: plannedCost,
    };
    try {
      await updateMutation.mutateAsync(payload);
      setDirty(false);
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const isPickup = shipment.delivery_type === "pickup";

  return (
    <section style={{ display: "grid", gap: "0.6rem" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <h3 style={{ margin: 0, fontSize: "1.05rem" }}>Рейс и документы</h3>
        {!readOnly && dirty && (
          <Button onClick={save} disabled={updateMutation.isPending}>
            {updateMutation.isPending ? "Сохранение..." : "Сохранить поля"}
          </Button>
        )}
      </div>

      {error && <Alert tone="error">{error}</Alert>}

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
          gap: "0.75rem",
        }}
      >
        {isPickup ? (
          <label style={labelStyle}>
            <span style={{ fontWeight: 600, color: "#344054" }}>№ доверенности</span>
            <Input
              type="text"
              value={draft.proxy_no}
              onChange={(e) => patch({ proxy_no: e.target.value })}
              disabled={readOnly}
            />
          </label>
        ) : (
          <label style={labelStyle}>
            <span style={{ fontWeight: 600, color: "#344054" }}>Перевозчик</span>
            <CarrierAutocomplete
              selected={draft.carrier}
              onSelect={(carrier) => patch({ carrier })}
              disabled={readOnly}
            />
          </label>
        )}
        <label style={labelStyle}>
          <span style={{ fontWeight: 600, color: "#344054" }}>Водитель</span>
          <Input
            type="text"
            value={draft.driver_name}
            onChange={(e) => patch({ driver_name: e.target.value })}
            disabled={readOnly}
          />
        </label>
        <label style={labelStyle}>
          <span style={{ fontWeight: 600, color: "#344054" }}>ТС (а/м, прицеп)</span>
          <Input
            type="text"
            value={draft.vehicle_text}
            onChange={(e) => patch({ vehicle_text: e.target.value })}
            disabled={readOnly}
          />
        </label>
        <label style={labelStyle}>
          <span style={{ fontWeight: 600, color: "#344054" }}>Класс ТС</span>
          <select
            value={draft.vehicle_class}
            onChange={(e) => patch({ vehicle_class: e.target.value as "" | VehicleClass })}
            disabled={readOnly}
            style={selectStyle}
          >
            <option value="">—</option>
            <option value="t20">до 19,8 т</option>
            <option value="t30plus">30 т+</option>
          </select>
        </label>
        <label style={labelStyle}>
          <span style={{ fontWeight: 600, color: "#344054" }}>УПД</span>
          <Input
            type="text"
            value={draft.upd_no}
            onChange={(e) => patch({ upd_no: e.target.value })}
            disabled={readOnly}
          />
        </label>
        <label style={labelStyle}>
          <span style={{ fontWeight: 600, color: "#344054" }}>№ заявки на фрахт</span>
          <Input
            type="text"
            value={draft.freight_request_no}
            onChange={(e) => patch({ freight_request_no: e.target.value })}
            disabled={readOnly}
          />
        </label>
        <label style={labelStyle}>
          <span style={{ fontWeight: 600, color: "#344054" }}>Стоимость план, ₽</span>
          <Input
            type="text"
            inputMode="decimal"
            value={draft.planned_cost}
            onChange={(e) => patch({ planned_cost: e.target.value })}
            disabled={readOnly}
          />
        </label>
      </div>

    </section>
  );
};
