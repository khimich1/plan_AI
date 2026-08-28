import { useMemo, useState, type CSSProperties } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Input } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import { currentMonthBounds } from "@/features/gsm/lib/fleetStatus";
import { formatLiters } from "@/features/gsm/lib/waybillWarnings";
import {
  useGsmTransactionsQuery,
  useGsmVehiclesQuery,
} from "@/features/gsm/hooks/useGsmQueries";
import type { TransactionListParams } from "@/features/gsm/types/gsm";

type Props = {
  onOpenCards?: () => void;
};

const wrap: CSSProperties = { display: "grid", gap: "1rem" };
const filters: CSSProperties = {
  display: "flex",
  gap: "0.75rem",
  flexWrap: "wrap",
  alignItems: "flex-end",
  border: "1px solid #eaecf0",
  borderRadius: 14,
  background: "#ffffff",
  padding: "0.9rem 1rem",
};
const label: CSSProperties = { display: "grid", gap: 4, fontSize: "0.85rem", color: "#475467" };
const selectStyle: CSSProperties = {
  border: "1px solid #d0d5dd",
  borderRadius: 12,
  padding: "0.8rem 0.9rem",
  background: "#ffffff",
  minWidth: 180,
};
const th: CSSProperties = {
  textAlign: "left",
  padding: "0.55rem 0.65rem",
  fontSize: "0.8rem",
  color: "#475467",
  borderBottom: "1px solid #eaecf0",
};
const td: CSSProperties = {
  padding: "0.55rem 0.65rem",
  fontSize: "0.9rem",
  borderBottom: "1px solid #f2f4f7",
};

const formatAmount = (value: number): string =>
  `${value.toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽`;

const serviceLabel = (type: string): string => {
  if (type === "fuel") return "Топливо";
  if (type === "wash") return "Мойка";
  return type;
};

export const TransactionsJournalView = ({ onOpenCards }: Props) => {
  const defaults = currentMonthBounds();
  const [periodFrom, setPeriodFrom] = useState(defaults.from);
  const [periodTo, setPeriodTo] = useState(defaults.to);
  const [vehicleId, setVehicleId] = useState<number | "">("");
  const [serviceType, setServiceType] = useState("");

  const params: TransactionListParams = useMemo(
    () => ({
      periodFrom,
      periodTo,
      vehicleId: vehicleId === "" ? undefined : vehicleId,
      serviceType: serviceType || undefined,
    }),
    [periodFrom, periodTo, vehicleId, serviceType],
  );

  const txQuery = useGsmTransactionsQuery(params);
  const vehiclesQuery = useGsmVehiclesQuery(true);
  const vehicles = vehiclesQuery.data ?? [];
  const vehiclesById = useMemo(() => {
    const map = new Map<number, string>();
    for (const v of vehicles) {
      map.set(v.id, `${v.name} (${v.plate_number})`);
    }
    return map;
  }, [vehicles]);

  const body = txQuery.data;

  return (
    <section style={wrap} aria-label="Журнал транзакций ГСМ">
      <div style={filters}>
        <label style={label}>
          Машина
          <select
            aria-label="Фильтр машины"
            value={vehicleId === "" ? "" : String(vehicleId)}
            onChange={(e) => setVehicleId(e.target.value === "" ? "" : Number(e.target.value))}
            style={selectStyle}
          >
            <option value="">Все</option>
            {vehicles.map((v) => (
              <option key={v.id} value={v.id}>
                {v.name}
              </option>
            ))}
          </select>
        </label>
        <label style={label}>
          Тип
          <select
            aria-label="Фильтр типа"
            value={serviceType}
            onChange={(e) => setServiceType(e.target.value)}
            style={selectStyle}
          >
            <option value="">Все</option>
            <option value="fuel">Топливо</option>
            <option value="wash">Мойка</option>
          </select>
        </label>
        <label style={label}>
          С
          <div style={{ width: 150 }}>
            <Input
              type="date"
              aria-label="Период с"
              value={periodFrom}
              onChange={(e) => setPeriodFrom(e.target.value)}
            />
          </div>
        </label>
        <label style={label}>
          По
          <div style={{ width: 150 }}>
            <Input
              type="date"
              aria-label="Период по"
              value={periodTo}
              onChange={(e) => setPeriodTo(e.target.value)}
            />
          </div>
        </label>
      </div>

      {txQuery.isLoading && <Spinner />}
      {txQuery.error && <Alert tone="error">{formatGsmError(txQuery.error)}</Alert>}

      {body && (
        <div style={{ overflowX: "auto", border: "1px solid #eaecf0", borderRadius: 14, background: "#fff" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={th}>Дата/время</th>
                <th style={th}>Машина</th>
                <th style={th}>Карта</th>
                <th style={th}>Услуга</th>
                <th style={th}>Литры</th>
                <th style={th}>Сумма</th>
                <th style={th}>АЗС</th>
              </tr>
            </thead>
            <tbody>
              {body.rows.map((row, index) => {
                const unbound = row.vehicle_id == null;
                return (
                  <tr
                    key={`${row.ts}-${row.card_number}-${index}`}
                    style={{ background: unbound ? "#fef3f2" : undefined }}
                    data-testid={unbound ? "unbound-card-row" : undefined}
                  >
                    <td style={td}>{row.ts}</td>
                    <td style={td}>
                      {row.vehicle_id == null ? (
                        <button
                          type="button"
                          onClick={onOpenCards}
                          style={{
                            border: "none",
                            background: "transparent",
                            color: "#b42318",
                            cursor: "pointer",
                            textDecoration: "underline",
                            padding: 0,
                            fontWeight: 600,
                          }}
                        >
                          Не привязана → Справочники → Карты
                        </button>
                      ) : (
                        vehiclesById.get(row.vehicle_id) ?? `#${row.vehicle_id}`
                      )}
                    </td>
                    <td style={td}>{row.card_number}</td>
                    <td style={td}>{serviceLabel(row.service_type)}</td>
                    <td style={td}>{formatLiters(row.qty_liters)}</td>
                    <td style={td}>{formatAmount(row.amount)}</td>
                    <td style={td}>{row.address ?? "—"}</td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr>
                <td style={{ ...td, fontWeight: 700 }} colSpan={4}>
                  Итого ({body.total_count})
                </td>
                <td style={{ ...td, fontWeight: 700 }} data-testid="tx-sum-liters">
                  {formatLiters(body.sum_liters)}
                </td>
                <td style={{ ...td, fontWeight: 700 }} data-testid="tx-sum-amount">
                  {formatAmount(body.sum_amount)}
                </td>
                <td style={td} />
              </tr>
            </tfoot>
          </table>
        </div>
      )}
    </section>
  );
};
