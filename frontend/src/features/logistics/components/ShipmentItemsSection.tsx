import { useEffect, useId, useMemo, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { getErrorMessage } from "@/shared/lib/apiError";
import {
  useConfirmItemsMutation,
  usePileCatalogQuery,
  useProposeMutation,
} from "@/features/logistics/hooks/useLogisticsQueries";
import { LayoutBlock } from "@/features/logistics/components/LayoutBlock";
import {
  VEHICLE_CLASS_LIMITS_FALLBACK_KG,
  formatWeightKg,
  vehicleClassLabel,
} from "@/features/logistics/lib/logisticsFormat";
import {
  applyPileCatalogEntry,
  draftFreeRow,
  draftFromAvailablePlate,
  draftFromProposed,
  draftFromSaved,
  draftItemLabel,
  draftRowWeightKg,
  draftTotalWeightKg,
  draftsToPayload,
  savedItemsVersion,
  type DraftItem,
} from "@/features/logistics/lib/draftItems";
import type {
  AvailablePlate,
  LayoutMetadata,
  OrderRemainderItem,
  ProposedItem,
  ProposeWarning,
  ShipmentDetails,
  VehicleClass,
} from "@/features/logistics/types/logistics";

type Props = {
  shipment: ShipmentDetails;
  readOnly: boolean;
};

const cellStyle: React.CSSProperties = { padding: "0.45rem", verticalAlign: "middle" };

const numericOrNull = (raw: string): number | null => {
  const value = Number(raw.replace(",", "."));
  return raw.trim() !== "" && Number.isFinite(value) ? value : null;
};

export const ShipmentItemsSection = ({ shipment, readOnly }: Props) => {
  const pileListId = useId();
  const [drafts, setDrafts] = useState<DraftItem[]>(() => shipment.items.map(draftFromSaved));
  const [dirty, setDirty] = useState(false);
  const [notFit, setNotFit] = useState<ProposedItem[]>([]);
  const [warnings, setWarnings] = useState<ProposeWarning[]>([]);
  const [orderRemainder, setOrderRemainder] = useState<OrderRemainderItem[]>([]);
  const [limits, setLimits] = useState<Record<VehicleClass, number> | null>(null);
  const [layout, setLayout] = useState<LayoutMetadata | null>(null);
  const [layoutWeight, setLayoutWeight] = useState<number | null>(null);
  const [layoutMaxWeight, setLayoutMaxWeight] = useState<number | null>(null);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [notFitExpanded, setNotFitExpanded] = useState(false);
  const [orderRemainderExpanded, setOrderRemainderExpanded] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);

  const proposeMutation = useProposeMutation(shipment.id);
  const confirmMutation = useConfirmItemsMutation(shipment.id);
  const pileCatalogQuery = usePileCatalogQuery("", !readOnly);
  const pileCatalog = pileCatalogQuery.data ?? [];

  const savedVersion = savedItemsVersion(shipment.items);
  useEffect(() => {
    setDrafts(shipment.items.map(draftFromSaved));
    setDirty(false);
    setNotFit([]);
    setWarnings([]);
    setOrderRemainder([]);
    setLayout(null);
    setLayoutWeight(null);
    setLayoutMaxWeight(null);
    setPickerOpen(false);
    setNotFitExpanded(false);
    setOrderRemainderExpanded(false);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [savedVersion, shipment.id]);

  const availableByPlateId = useMemo(() => {
    const map = new Map<number, { plate: AvailablePlate; kpId: number }>();
    for (const group of shipment.available_by_kp) {
      for (const plate of group.plates) {
        map.set(plate.completed_plate_id, { plate, kpId: group.kp_id });
      }
    }
    return map;
  }, [shipment.available_by_kp]);

  const totalWeight = draftTotalWeightKg(drafts);
  const vehicleClass = shipment.vehicle_class;
  const classLimit = vehicleClass
    ? (limits?.[vehicleClass] ?? VEHICLE_CLASS_LIMITS_FALLBACK_KG[vehicleClass])
    : null;
  const overload = classLimit != null && totalWeight > classLimit;

  const clearLayout = () => {
    setLayout(null);
    setLayoutWeight(null);
    setLayoutMaxWeight(null);
  };

  const mutateDrafts = (updater: (prev: DraftItem[]) => DraftItem[]) => {
    setDrafts((prev) => updater(prev));
    setDirty(true);
    clearLayout();
  };

  const patchDraft = (key: string, partial: Partial<DraftItem>) =>
    mutateDrafts((prev) => prev.map((item) => (item.key === key ? { ...item, ...partial } : item)));

  const runPropose = async () => {
    setError(null);
    setInfo(null);
    try {
      const response = await proposeMutation.mutateAsync(shipment.vehicle_class);
      setDrafts(response.items.map(draftFromProposed));
      setNotFit(response.not_fit);
      setWarnings(response.warnings ?? []);
      setOrderRemainder(response.order_remainder ?? []);
      setLimits(response.vehicle_class_limits_kg);
      const layoutClass = (response.vehicle_class ?? null) as VehicleClass | null;
      setLayout(response.layout ?? null);
      setLayoutWeight(response.layout ? response.total_weight_kg : null);
      setLayoutMaxWeight(
        response.layout && layoutClass
          ? (response.vehicle_class_limits_kg?.[layoutClass] ?? null)
          : null,
      );
      setDirty(true);
      setInfo(
        response.items.length > 0
          ? "Предложенный состав подставлен в редактор — проверьте и утвердите."
          : "Свободных плит по заказам рейса на СГП не найдено.",
      );
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const runConfirm = async () => {
    setError(null);
    setInfo(null);
    const invalid = drafts.some((item) => !Number.isInteger(item.qty) || item.qty < 1);
    if (invalid) {
      setError("Количество в каждой строке должно быть целым числом не меньше 1.");
      return;
    }
    const emptyMark = drafts.some((item) => item.item_type === "free" && !item.mark.trim());
    if (emptyMark) {
      setError("У свободной строки заполните марку (например, С60.30).");
      return;
    }
    try {
      await confirmMutation.mutateAsync(draftsToPayload(drafts));
      setDirty(false);
      setNotFit([]);
      setInfo("Состав утверждён и сохранён.");
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const busy = proposeMutation.isPending || confirmMutation.isPending;

  if (readOnly) {
    return (
      <section style={{ display: "grid", gap: "0.6rem" }}>
        <h3 style={{ margin: 0, fontSize: "1.05rem" }}>Состав рейса</h3>
        {shipment.items.length === 0 ? (
          <Alert tone="info">Состав не задан.</Alert>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.9rem" }}>
              <thead>
                <tr style={{ textAlign: "left", borderBottom: "1px solid #eaecf0" }}>
                  <th style={cellStyle}>№</th>
                  <th style={cellStyle}>Позиция</th>
                  <th style={cellStyle}>Размер</th>
                  <th style={cellStyle}>Кол-во</th>
                  <th style={cellStyle}>Вес, кг</th>
                  <th style={cellStyle}>Заметка</th>
                </tr>
              </thead>
              <tbody>
                {shipment.items.map((item, index) => (
                  <tr key={item.id} style={{ borderBottom: "1px solid #f2f4f7" }}>
                    <td style={cellStyle}>{index + 1}</td>
                    <td style={{ ...cellStyle, fontWeight: 600 }}>
                      {item.item_type === "plate" ? item.plate_name : item.mark || "Свободная строка"}
                    </td>
                    <td style={cellStyle}>
                      {item.item_type === "plate"
                        ? `${item.length_m ?? "—"}×${item.width_m ?? "—"} / ${item.load_class ?? "—"}`
                        : "—"}
                    </td>
                    <td style={cellStyle}>{item.qty}</td>
                    <td style={cellStyle}>{formatWeightKg(item.weight_kg)}</td>
                    <td style={cellStyle}>{item.note || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
        <div style={{ color: "#475467", fontSize: "0.9rem" }}>
          Итого: {formatWeightKg(shipment.total_weight_kg)}
        </div>
      </section>
    );
  }

  return (
    <section style={{ display: "grid", gap: "0.6rem" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "0.5rem" }}>
        <h3 style={{ margin: 0, fontSize: "1.05rem" }}>Состав рейса</h3>
        <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
          <Button variant="secondary" onClick={runPropose} disabled={busy}>
            {proposeMutation.isPending ? "Подбор..." : "Предложить состав"}
          </Button>
          {dirty && (
            <>
              <Button variant="ghost" onClick={() => { setDrafts(shipment.items.map(draftFromSaved)); setDirty(false); setNotFit([]); clearLayout(); }} disabled={busy}>
                Сбросить правки
              </Button>
              <Button onClick={runConfirm} disabled={busy}>
                {confirmMutation.isPending ? "Сохранение..." : "Утвердить состав"}
              </Button>
            </>
          )}
        </div>
      </div>

      {error && <Alert tone="error">{error}</Alert>}
      {info && <Alert tone="info">{info}</Alert>}

      {drafts.length === 0 ? (
        <Alert tone="info">
          Состав пуст. Нажмите «Предложить состав» или добавьте строки вручную.
        </Alert>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.9rem" }}>
            <thead>
              <tr style={{ textAlign: "left", borderBottom: "1px solid #eaecf0" }}>
                <th style={cellStyle}>№</th>
                <th style={cellStyle}>Позиция</th>
                <th style={cellStyle}>Размер</th>
                <th style={cellStyle}>КП</th>
                <th style={cellStyle}>Кол-во</th>
                <th style={cellStyle}>Вес, кг</th>
                <th style={cellStyle}>Заметка</th>
                <th style={cellStyle} aria-label="Удалить" />
              </tr>
            </thead>
            <tbody>
              {drafts.map((item, index) => {
                const available = item.completed_plate_id != null
                  ? availableByPlateId.get(item.completed_plate_id)?.plate.available_qty
                  : undefined;
                return (
                  <tr key={item.key} style={{ borderBottom: "1px solid #f2f4f7" }}>
                    <td style={cellStyle}>{index + 1}</td>
                    <td style={{ ...cellStyle, minWidth: 180 }}>
                      {item.item_type === "plate" ? (
                        <span style={{ fontWeight: 600 }}>{draftItemLabel(item)}</span>
                      ) : (
                        <>
                          <Input
                            type="text"
                            value={item.mark}
                            placeholder="Марка (С60.30)"
                            list={pileListId}
                            onChange={(e) => {
                              const mark = e.target.value;
                              const entry = pileCatalog.find(
                                (p) => p.mark.trim().toLowerCase() === mark.trim().toLowerCase(),
                              );
                              if (entry) {
                                patchDraft(item.key, applyPileCatalogEntry(item, entry));
                              } else {
                                patchDraft(item.key, { mark });
                              }
                            }}
                          />
                          {item.unit_weight_kg != null && (
                            <div style={{ fontSize: "0.8rem", color: "#667085", marginTop: 2 }}>
                              автовес: {formatWeightKg(item.unit_weight_kg)} / шт
                            </div>
                          )}
                        </>
                      )}
                    </td>
                    <td style={cellStyle}>{item.item_type === "plate" ? item.dims || "—" : "—"}</td>
                    <td style={cellStyle}>
                      {item.item_type === "free" ? (
                        <select
                          value={item.kp_id ?? ""}
                          onChange={(e) =>
                            patchDraft(item.key, {
                              kp_id: e.target.value ? Number(e.target.value) : null,
                            })
                          }
                          style={{
                            border: "1px solid #d0d5dd",
                            borderRadius: 10,
                            padding: "0.45rem 0.5rem",
                            background: "#ffffff",
                          }}
                        >
                          <option value="">—</option>
                          {shipment.orders.map((order) => (
                            <option key={order.kp_id} value={order.kp_id}>
                              КП №{order.kp_id}
                            </option>
                          ))}
                        </select>
                      ) : (
                        item.kp_id != null ? `№${item.kp_id}` : "—"
                      )}
                    </td>
                    <td style={{ ...cellStyle, width: 90 }}>
                      <Input
                        type="number"
                        min={1}
                        value={item.qty}
                        onChange={(e) =>
                          patchDraft(item.key, {
                            qty: Math.max(1, Math.trunc(Number(e.target.value) || 1)),
                            // Автовес пересчитываем от unit×qty; ручной вес свободной строки не трогаем
                            ...(item.weight_manual ? {} : { weight_kg: null }),
                          })
                        }
                      />
                      {available != null && (
                        <div style={{ fontSize: "0.8rem", color: "#667085", marginTop: 2 }}>
                          свободно: {available}
                        </div>
                      )}
                    </td>
                    <td style={{ ...cellStyle, width: 130 }}>
                      {item.item_type === "free" ? (
                        <>
                          <Input
                            type="text"
                            inputMode="decimal"
                            value={item.weight_manual && item.weight_kg != null ? String(item.weight_kg) : ""}
                            placeholder={
                              item.unit_weight_kg != null
                                ? `авто: ${Math.round(item.unit_weight_kg * item.qty).toLocaleString("ru-RU")}`
                                : "вес, кг"
                            }
                            onChange={(e) => {
                              const value = numericOrNull(e.target.value);
                              patchDraft(item.key, {
                                weight_manual: value != null,
                                weight_kg: value,
                              });
                            }}
                          />
                          {!item.weight_manual && draftRowWeightKg(item) != null && (
                            <div style={{ fontSize: "0.8rem", color: "#667085", marginTop: 2 }}>
                              = {formatWeightKg(draftRowWeightKg(item))}
                            </div>
                          )}
                        </>
                      ) : (
                        formatWeightKg(draftRowWeightKg(item))
                      )}
                    </td>
                    <td style={{ ...cellStyle, minWidth: 140 }}>
                      <Input
                        type="text"
                        value={item.note}
                        placeholder="Заметка укладки"
                        onChange={(e) => patchDraft(item.key, { note: e.target.value })}
                      />
                    </td>
                    <td style={cellStyle}>
                      <Button
                        variant="ghost"
                        aria-label={`Удалить строку ${index + 1}`}
                        onClick={() => mutateDrafts((prev) => prev.filter((p) => p.key !== item.key))}
                        disabled={busy}
                      >
                        ×
                      </Button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      <LayoutBlock layout={layout} totalWeightKg={layoutWeight} maxWeightKg={layoutMaxWeight} />

      <datalist id={pileListId}>
        {pileCatalog.map((entry) => (
          <option key={entry.id} value={entry.mark}>
            {`${entry.mark} — ${formatWeightKg(entry.weight_kg)}`}
          </option>
        ))}
      </datalist>

      <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
        <Button variant="secondary" onClick={() => setPickerOpen((v) => !v)} disabled={busy}>
          {pickerOpen ? "Скрыть доступные плиты" : "+ Плита со СГП"}
        </Button>
        <Button variant="secondary" onClick={() => mutateDrafts((prev) => [...prev, draftFreeRow(shipment.orders[0]?.kp_id ?? null)])} disabled={busy}>
          + Свободная строка
        </Button>
      </div>

      {pickerOpen && (
        <div
          style={{
            border: "1px solid #eaecf0",
            borderRadius: 12,
            padding: "0.75rem",
            display: "grid",
            gap: "0.6rem",
            background: "#fcfcfd",
          }}
        >
          {shipment.available_by_kp.length === 0 && (
            <span style={{ color: "#667085" }}>Свободных плит по заказам рейса нет.</span>
          )}
          {shipment.available_by_kp.map((group) => (
            <div key={group.kp_id} style={{ display: "grid", gap: "0.35rem" }}>
              <strong>КП №{group.kp_id}</strong>
              {group.plates.length === 0 ? (
                <span style={{ color: "#667085" }}>Нет свободных плит.</span>
              ) : (
                group.plates.map((plate) => (
                  <div
                    key={plate.completed_plate_id}
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      gap: "0.5rem",
                      border: "1px solid #f2f4f7",
                      borderRadius: 10,
                      padding: "0.4rem 0.6rem",
                      background: "#ffffff",
                    }}
                  >
                    <span>
                      {plate.plate_name} · {plate.length_m ?? "—"}×{plate.width_m ?? "—"} /{" "}
                      {plate.load_class ?? "—"} · свободно {plate.available_qty} шт
                      {plate.unit_weight_kg != null && ` · ${formatWeightKg(plate.unit_weight_kg)}/шт`}
                    </span>
                    <Button
                      variant="secondary"
                      onClick={() => mutateDrafts((prev) => [...prev, draftFromAvailablePlate(plate, group.kp_id)])}
                      disabled={busy}
                    >
                      Добавить
                    </Button>
                  </div>
                ))
              )}
            </div>
          ))}
        </div>
      )}

      {vehicleClass === "t30plus" && (
        <Alert tone="info">
          Для класса «30 т+» автоподбор идёт по весу (без правил укладки). Для подбора по геометрии
          выберите класс «до 19,8 т».
        </Alert>
      )}

      {warnings.length > 0 && (
        <div
          style={{
            border: "1px solid #b2ddff",
            borderRadius: 12,
            background: "#eff8ff",
            padding: "0.75rem",
            display: "grid",
            gap: "0.35rem",
          }}
        >
          <strong style={{ color: "#175cd3" }}>Предупреждения:</strong>
          {warnings.map((warning, index) => (
            <span key={index} style={{ color: "#1849a9", fontSize: "0.9rem" }}>
              {warning.message}
              {warning.kp_ids && warning.kp_ids.length > 0
                ? ` (КП №${warning.kp_ids.join(", №")})`
                : ""}
            </span>
          ))}
        </div>
      )}

      {orderRemainder.length > 0 && (
        <div
          style={{
            border: "1px solid #eaecf0",
            borderRadius: 12,
            background: "#f9fafb",
            padding: "0.75rem",
            display: "grid",
            gap: "0.35rem",
          }}
        >
          <button
            type="button"
            onClick={() => setOrderRemainderExpanded((expanded) => !expanded)}
            aria-expanded={orderRemainderExpanded}
            style={{
              display: "flex",
              alignItems: "center",
              gap: "0.35rem",
              width: "100%",
              border: "none",
              background: "transparent",
              padding: 0,
              cursor: "pointer",
              textAlign: "left",
              color: "#344054",
              fontWeight: 700,
              font: "inherit",
            }}
          >
            <span aria-hidden="true">{orderRemainderExpanded ? "▾" : "▸"}</span>
            <span>
              {orderRemainderExpanded
                ? "Остаток по заказу (на следующий рейс)"
                : `Остаток по заказу (на следующий рейс) (${orderRemainder.length} поз.)`}
            </span>
          </button>
          {orderRemainderExpanded &&
            orderRemainder.map((line) => (
              <span key={line.completed_plate_id} style={{ color: "#475467", fontSize: "0.9rem" }}>
                {line.plate_name} · {line.qty_remaining} шт (КП №{line.kp_id})
              </span>
            ))}
        </div>
      )}

      {notFit.length > 0 && (
        <div
          style={{
            border: "1px solid #fedf89",
            borderRadius: 12,
            background: "#fffaeb",
            padding: "0.75rem",
            display: "grid",
            gap: "0.4rem",
          }}
        >
          <button
            type="button"
            onClick={() => setNotFitExpanded((expanded) => !expanded)}
            aria-expanded={notFitExpanded}
            style={{
              display: "flex",
              alignItems: "center",
              gap: "0.35rem",
              width: "100%",
              border: "none",
              background: "transparent",
              padding: 0,
              cursor: "pointer",
              textAlign: "left",
              color: "#b54708",
              fontWeight: 700,
              font: "inherit",
            }}
          >
            <span aria-hidden="true">{notFitExpanded ? "▾" : "▸"}</span>
            <span>
              {notFitExpanded
                ? "Не влезло в лимит класса ТС"
                : `Не влезло в лимит класса ТС (${notFit.length} поз.)`}
            </span>
          </button>
          {notFitExpanded &&
            notFit.map((item, index) => (
              <div
                key={index}
                style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "0.5rem" }}
              >
                <span style={{ color: "#7a4a06" }}>
                  {item.plate_name} · {item.qty} шт · {formatWeightKg(item.weight_kg)}
                  {item.reason_text ? ` — ${item.reason_text}` : ""}
                </span>
                <Button
                  variant="secondary"
                  onClick={() => mutateDrafts((prev) => [...prev, draftFromProposed(item)])}
                  disabled={busy}
                >
                  Добавить всё равно
                </Button>
              </div>
            ))}
        </div>
      )}

      <div
        style={{
          display: "flex",
          gap: "1rem",
          alignItems: "center",
          flexWrap: "wrap",
          borderTop: "1px solid #eaecf0",
          paddingTop: "0.6rem",
        }}
      >
        <strong>Σ веса: {formatWeightKg(totalWeight)}</strong>
        {vehicleClass && classLimit != null && (
          <span style={{ color: "#475467" }}>
            Лимит ТС ({vehicleClassLabel(vehicleClass)}): {formatWeightKg(classLimit)}
          </span>
        )}
        {overload && (
          <span role="alert" style={{ color: "#b42318", fontWeight: 700 }}>
            Перегруз: состав тяжелее лимита на {formatWeightKg(totalWeight - classLimit)}
          </span>
        )}
        {!vehicleClass && (
          <span style={{ color: "#667085", fontSize: "0.9rem" }}>
            Укажите класс ТС в полях рейса, чтобы видеть лимит веса.
          </span>
        )}
      </div>
    </section>
  );
};
