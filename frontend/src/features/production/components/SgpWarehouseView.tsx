import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Modal } from "@/shared/ui/Modal";
import { Spinner } from "@/shared/ui/Spinner";
import { getErrorMessage } from "@/shared/lib/apiError";
import type { SgpFilter, SgpPlateItem } from "@/features/production/api/sgpApi";
import {
  getUnlinkConfirmButtonLabel,
  getUnlinkPrompt,
  isSingleQtyUnlink,
  resolveUnlinkSubmitQty,
} from "@/features/production/lib/sgpWarehouseUnlink";
import {
  useSgpPlatesQuery,
  useSgpRelinkMutation,
  useSgpUnlinkMutation,
} from "@/features/production/hooks/useSgpQueries";

const FILTERS: { value: SgpFilter; label: string }[] = [
  { value: "all", label: "Все" },
  { value: "linked", label: "С КП" },
  { value: "unlinked", label: "Без КП" },
];

const COL_COUNT = 9;

const formatDims = (p: SgpPlateItem): string => {
  const L = p.length_m != null ? `${p.length_m}` : "—";
  const W = p.width_m != null ? `${p.width_m}` : "—";
  const load = p.load_class != null ? String(p.load_class) : "—";
  return `${L}×${W} / ${load}`;
};

export const SgpWarehouseView = () => {
  const [filter, setFilter] = useState<SgpFilter>("all");
  const query = useSgpPlatesQuery(filter);
  const unlinkMutation = useSgpUnlinkMutation();
  const relinkMutation = useSgpRelinkMutation();

  const [unlinkTarget, setUnlinkTarget] = useState<SgpPlateItem | null>(null);
  const [unlinkQty, setUnlinkQty] = useState(1);
  const [relinkTarget, setRelinkTarget] = useState<SgpPlateItem | null>(null);
  const [relinkKpId, setRelinkKpId] = useState("");
  const [relinkQty, setRelinkQty] = useState(1);
  const [actionMessage, setActionMessage] = useState<string | null>(null);
  const [actionError, setActionError] = useState<string | null>(null);

  const activeRowRef = useRef<HTMLTableRowElement>(null);

  const items = query.data?.items ?? [];
  const busy = unlinkMutation.isPending || relinkMutation.isPending;

  const summary = useMemo(() => {
    const linked = items.filter((i) => i.kp_id != null).length;
    const free = items.length - linked;
    return { linked, free, total: items.length };
  }, [items]);

  useEffect(() => {
    if (!unlinkTarget && !relinkTarget) {
      return;
    }
    activeRowRef.current?.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }, [unlinkTarget, relinkTarget]);

  const openUnlink = (plate: SgpPlateItem) => {
    setRelinkTarget(null);
    setActionError(null);
    setActionMessage(null);
    if (unlinkTarget?.id === plate.id) {
      setUnlinkTarget(null);
      return;
    }
    setUnlinkTarget(plate);
    setUnlinkQty(plate.qty > 0 ? Math.min(1, plate.qty) : 1);
  };

  const openRelink = (plate: SgpPlateItem) => {
    setUnlinkTarget(null);
    setActionError(null);
    setActionMessage(null);
    setRelinkTarget(plate);
    setRelinkQty(plate.qty > 0 ? Math.min(1, plate.qty) : 1);
    setRelinkKpId("");
  };

  const closeUnlink = () => setUnlinkTarget(null);

  const closeRelink = () => {
    setRelinkTarget(null);
    setActionError(null);
  };

  const submitUnlink = async () => {
    if (!unlinkTarget) return;
    const qty = resolveUnlinkSubmitQty(unlinkTarget, unlinkQty);
    try {
      const res = await unlinkMutation.mutateAsync({
        sgpId: unlinkTarget.id,
        qty,
      });
      setActionMessage(res.message || "Отвязано");
      setUnlinkTarget(null);
    } catch (err) {
      setActionError(getErrorMessage(err));
    }
  };

  const submitRelink = async () => {
    if (!relinkTarget) return;
    const targetKpId = Number(relinkKpId);
    if (!Number.isFinite(targetKpId) || targetKpId < 1) {
      setActionError("Укажите номер целевого КП");
      return;
    }
    try {
      const res = await relinkMutation.mutateAsync({
        sgpId: relinkTarget.id,
        targetKpId,
        qty: relinkQty,
      });
      setActionMessage(res.message || "Перепривязано");
      setRelinkTarget(null);
    } catch (err) {
      setActionError(getErrorMessage(err));
    }
  };

  const renderUnlinkPanel = (plate: SgpPlateItem) => {
    const isSingle = isSingleQtyUnlink(plate.qty);

    return (
      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: "0.75rem",
          alignItems: "center",
        }}
      >
        <span style={{ color: "#344054", fontWeight: 600 }}>{getUnlinkPrompt(plate)}</span>
        {!isSingle && (
          <label
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              color: "#475467",
            }}
          >
            Количество (макс. {plate.qty})
            <input
              type="number"
              min={1}
              max={plate.qty}
              value={unlinkQty}
              onChange={(e) =>
                setUnlinkQty(
                  Math.max(1, Math.min(plate.qty, Number(e.target.value) || 1)),
                )
              }
              style={{ width: 72 }}
            />
          </label>
        )}
        <div style={{ display: "flex", gap: 8 }}>
          <Button onClick={submitUnlink} disabled={busy}>
            {getUnlinkConfirmButtonLabel(plate.qty)}
          </Button>
          <Button variant="secondary" onClick={closeUnlink} disabled={busy}>
            Отмена
          </Button>
        </div>
      </div>
    );
  };

  return (
    <section style={{ display: "grid", gap: "1rem" }}>
      <header>
        <h2 style={{ margin: 0, fontSize: "1.25rem" }}>Склад готовой продукции</h2>
        <p style={{ margin: "0.35rem 0 0", color: "#475467" }}>
          Физические плиты на СГП. Отвязка возвращает потребность в КП; перепривязка —
          только при точном совпадении номенклатуры.
        </p>
      </header>

      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
        {FILTERS.map((f) => (
          <button
            key={f.value}
            type="button"
            onClick={() => setFilter(f.value)}
            style={{
              border: "1px solid #d0d5dd",
              borderRadius: 8,
              padding: "0.4rem 0.75rem",
              cursor: "pointer",
              background: filter === f.value ? "#1d4ed8" : "#fff",
              color: filter === f.value ? "#fff" : "#344054",
              fontWeight: 600,
            }}
          >
            {f.label}
          </button>
        ))}
        <span style={{ color: "#667085", fontSize: "0.9rem", marginLeft: 8 }}>
          Всего: {summary.total} · с КП: {summary.linked} · свободно: {summary.free}
        </span>
      </div>

      {actionMessage && <Alert tone="success">{actionMessage}</Alert>}
      {actionError && !relinkTarget && <Alert tone="error">{actionError}</Alert>}
      {query.isError && (
        <Alert tone="error">{getErrorMessage(query.error)}</Alert>
      )}

      {query.isLoading ? (
        <Spinner />
      ) : items.length === 0 ? (
        <Alert tone="info">На складе пока нет плит по выбранному фильтру.</Alert>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table
            style={{
              width: "100%",
              borderCollapse: "collapse",
              fontSize: "0.92rem",
            }}
          >
            <thead>
              <tr style={{ textAlign: "left", borderBottom: "1px solid #eaecf0" }}>
                <th style={{ padding: "0.55rem" }}>Плита</th>
                <th style={{ padding: "0.55rem" }}>Размер</th>
                <th style={{ padding: "0.55rem" }}>КП</th>
                <th style={{ padding: "0.55rem" }}>Заказчик</th>
                <th style={{ padding: "0.55rem" }}>Срок</th>
                <th style={{ padding: "0.55rem" }}>Qty</th>
                <th style={{ padding: "0.55rem" }}>N/M</th>
                <th style={{ padding: "0.55rem" }}>Дата</th>
                <th style={{ padding: "0.55rem" }}>Действия</th>
              </tr>
            </thead>
            <tbody>
              {items.map((plate) => {
                const isUnlinkOpen = unlinkTarget?.id === plate.id;
                const isRelinkActive = relinkTarget?.id === plate.id;
                const isActive = isUnlinkOpen || isRelinkActive;

                return (
                  <Fragment key={plate.id}>
                    <tr
                      ref={isActive ? activeRowRef : undefined}
                      style={{
                        borderBottom: isUnlinkOpen ? undefined : "1px solid #f2f4f7",
                        background: isActive ? "#eff6ff" : undefined,
                      }}
                    >
                      <td style={{ padding: "0.55rem", fontWeight: 600 }}>
                        {plate.plate_name}
                      </td>
                      <td style={{ padding: "0.55rem" }}>{formatDims(plate)}</td>
                      <td style={{ padding: "0.55rem" }}>
                        {plate.kp_id != null ? `#${plate.kp_id}` : "—"}
                      </td>
                      <td style={{ padding: "0.55rem" }}>
                        {plate.customer_name || "—"}
                      </td>
                      <td style={{ padding: "0.55rem" }}>
                        {plate.execution_terms || "—"}
                      </td>
                      <td style={{ padding: "0.55rem" }}>{plate.qty}</td>
                      <td style={{ padding: "0.55rem" }}>
                        {plate.sgp_progress
                          ? `${plate.sgp_progress.n}/${plate.sgp_progress.m}`
                          : "—"}
                      </td>
                      <td style={{ padding: "0.55rem" }}>
                        {plate.completed_date || "—"}
                      </td>
                      <td style={{ padding: "0.55rem" }}>
                        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {plate.kp_id != null && (
                            <Button
                              variant={isUnlinkOpen ? "primary" : "secondary"}
                              disabled={busy}
                              onClick={() => openUnlink(plate)}
                            >
                              {isUnlinkOpen ? "Отмена" : "Отвязать"}
                            </Button>
                          )}
                          {plate.kp_id == null && (
                            <Button
                              variant="secondary"
                              disabled={busy}
                              onClick={() => openRelink(plate)}
                            >
                              Перепривязать
                            </Button>
                          )}
                        </div>
                      </td>
                    </tr>
                    {isUnlinkOpen && (
                      <tr style={{ borderBottom: "1px solid #f2f4f7", background: "#f0f9ff" }}>
                        <td colSpan={COL_COUNT} style={{ padding: "0.65rem 0.75rem" }}>
                          {renderUnlinkPanel(plate)}
                        </td>
                      </tr>
                    )}
                  </Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      <Modal
        open={relinkTarget != null}
        onClose={closeRelink}
        title={
          relinkTarget
            ? `Перепривязать: ${relinkTarget.plate_name}`
            : "Перепривязать"
        }
        maxWidth={480}
      >
        {relinkTarget && (
          <div style={{ display: "grid", gap: "0.75rem" }}>
            <p style={{ margin: 0, color: "#475467" }}>
              Только при точном совпадении номенклатуры и открытой потребности у целевого КП.
            </p>
            <label style={{ display: "grid", gap: 4 }}>
              Целевой КП (номер)
              <input
                type="number"
                min={1}
                value={relinkKpId}
                onChange={(e) => setRelinkKpId(e.target.value)}
                placeholder="например 42"
                autoFocus
              />
            </label>
            <label style={{ display: "grid", gap: 4 }}>
              Количество (макс. {relinkTarget.qty})
              <input
                type="number"
                min={1}
                max={relinkTarget.qty}
                value={relinkQty}
                onChange={(e) =>
                  setRelinkQty(
                    Math.max(1, Math.min(relinkTarget.qty, Number(e.target.value) || 1)),
                  )
                }
              />
            </label>
            {actionError && relinkTarget && (
              <Alert tone="error">{actionError}</Alert>
            )}
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
              <Button variant="secondary" onClick={closeRelink} disabled={busy}>
                Отмена
              </Button>
              <Button onClick={submitRelink} disabled={busy}>
                Подтвердить
              </Button>
            </div>
          </div>
        )}
      </Modal>
    </section>
  );
};
