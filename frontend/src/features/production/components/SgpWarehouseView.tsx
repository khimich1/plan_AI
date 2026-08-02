import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Modal } from "@/shared/ui/Modal";
import { Spinner } from "@/shared/ui/Spinner";
import { getErrorMessage } from "@/shared/lib/apiError";
import type { SgpFilter, SgpPlateItem } from "@/features/production/api/sgpApi";
import {
  allGroupKeys,
  groupKeyForPlate,
  groupLabel,
  groupPlatesByKp,
  type SgpGroupKey,
  type SgpPlateGroup,
} from "@/features/production/lib/sgpWarehouseGroups";
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

const OUTER_COL_COUNT = 8;
const PLATE_COL_COUNT = 5;

const cellPad = { padding: "0.55rem" } as const;

const chevronButtonStyle = {
  border: "none",
  background: "transparent",
  cursor: "pointer",
  fontSize: "0.95rem",
  color: "#475467",
  padding: "0.1rem 0.25rem",
  lineHeight: 1,
} as const;

const formatDims = (p: SgpPlateItem): string => {
  const L = p.length_m != null ? `${p.length_m}` : "—";
  const W = p.width_m != null ? `${p.width_m}` : "—";
  const load = p.load_class != null ? String(p.load_class) : "—";
  return `${L}×${W} / ${load}`;
};

const formatProgress = (group: SgpPlateGroup): string => {
  if (!group.sgpProgress) return "—";
  return `${group.sgpProgress.n}/${group.sgpProgress.m}`;
};

export const SgpWarehouseView = () => {
  const [filter, setFilter] = useState<SgpFilter>("all");
  const query = useSgpPlatesQuery(filter);
  const unlinkMutation = useSgpUnlinkMutation();
  const relinkMutation = useSgpRelinkMutation();

  const [expandedGroupKeys, setExpandedGroupKeys] = useState<Set<SgpGroupKey>>(new Set());
  const [unlinkTarget, setUnlinkTarget] = useState<SgpPlateItem | null>(null);
  const [unlinkQty, setUnlinkQty] = useState(1);
  const [relinkTarget, setRelinkTarget] = useState<SgpPlateItem | null>(null);
  const [relinkKpId, setRelinkKpId] = useState("");
  const [relinkQty, setRelinkQty] = useState(1);
  const [actionMessage, setActionMessage] = useState<string | null>(null);
  const [actionError, setActionError] = useState<string | null>(null);

  const activeRowRef = useRef<HTMLTableRowElement>(null);

  const items = query.data?.items ?? [];
  const groups = useMemo(() => groupPlatesByKp(items), [items]);
  const busy = unlinkMutation.isPending || relinkMutation.isPending;

  const summary = useMemo(() => {
    const linked = items.filter((i) => i.kp_id != null).length;
    const free = items.length - linked;
    return { linked, free, total: items.length };
  }, [items]);

  const allExpanded =
    groups.length > 0 && groups.every((g) => expandedGroupKeys.has(g.key));

  useEffect(() => {
    setExpandedGroupKeys(new Set());
    setUnlinkTarget(null);
    setRelinkTarget(null);
  }, [filter]);

  useEffect(() => {
    if (!unlinkTarget && !relinkTarget) {
      return;
    }
    activeRowRef.current?.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }, [unlinkTarget, relinkTarget]);

  useEffect(() => {
    if (!unlinkTarget) return;
    const key = groupKeyForPlate(unlinkTarget);
    setExpandedGroupKeys((prev) => {
      if (prev.has(key)) return prev;
      const next = new Set(prev);
      next.add(key);
      return next;
    });
  }, [unlinkTarget]);

  const toggleGroup = (key: SgpGroupKey) => {
    setExpandedGroupKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  };

  const expandAll = () => {
    setExpandedGroupKeys(new Set(allGroupKeys(groups)));
  };

  const collapseAll = () => {
    setExpandedGroupKeys(new Set());
    setUnlinkTarget(null);
  };

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
    setExpandedGroupKeys((prev) => {
      const key = groupKeyForPlate(plate);
      if (prev.has(key)) return prev;
      const next = new Set(prev);
      next.add(key);
      return next;
    });
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

  const renderPlateRows = (group: SgpPlateGroup) => (
    <table
      style={{
        width: "100%",
        borderCollapse: "collapse",
        fontSize: "0.9rem",
      }}
    >
      <thead>
        <tr style={{ textAlign: "left", color: "#667085" }}>
          <th style={cellPad}>Плита</th>
          <th style={cellPad}>Размер</th>
          <th style={cellPad}>Qty</th>
          <th style={cellPad}>Дата</th>
          <th style={cellPad}>Действия</th>
        </tr>
      </thead>
      <tbody>
        {group.plates.map((plate) => {
          const isUnlinkOpen = unlinkTarget?.id === plate.id;
          const isRelinkActive = relinkTarget?.id === plate.id;
          const isActive = isUnlinkOpen || isRelinkActive;

          return (
            <Fragment key={plate.id}>
              <tr
                ref={isActive ? activeRowRef : undefined}
                style={{
                  borderTop: "1px solid #eef2f6",
                  background: isActive ? "#eff6ff" : undefined,
                }}
              >
                <td style={{ ...cellPad, fontWeight: 600 }}>{plate.plate_name}</td>
                <td style={cellPad}>{formatDims(plate)}</td>
                <td style={cellPad}>{plate.qty}</td>
                <td style={cellPad}>{plate.completed_date || "—"}</td>
                <td style={cellPad}>
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
                <tr style={{ background: "#f0f9ff" }}>
                  <td colSpan={PLATE_COL_COUNT} style={{ padding: "0.65rem 0.75rem" }}>
                    {renderUnlinkPanel(plate)}
                  </td>
                </tr>
              )}
            </Fragment>
          );
        })}
      </tbody>
    </table>
  );

  const renderGroupHeader = (group: SgpPlateGroup) => {
    const isExpanded = expandedGroupKeys.has(group.key);
    const label = groupLabel(group);

    return (
      <tr
        style={{
          borderTop: "1px solid #e4e7ec",
          background: "#f8fafc",
        }}
      >
        <td style={{ ...cellPad, width: 36, textAlign: "center" }}>
          <button
            type="button"
            onClick={() => toggleGroup(group.key)}
            aria-label={isExpanded ? `Свернуть ${label}` : `Развернуть ${label}`}
            aria-expanded={isExpanded}
            style={chevronButtonStyle}
          >
            {isExpanded ? "▾" : "▸"}
          </button>
        </td>
        <td style={{ ...cellPad, fontWeight: 700 }}>{label}</td>
        <td style={cellPad}>{group.customerName || "—"}</td>
        <td style={cellPad}>{group.executionTerms || "—"}</td>
        <td style={cellPad}>{group.positionCount}</td>
        <td style={cellPad}>{group.totalQty}</td>
        <td style={cellPad}>{formatProgress(group)}</td>
        <td style={cellPad} />
      </tr>
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
        {groups.length > 0 && (
          <Button variant="secondary" onClick={allExpanded ? collapseAll : expandAll}>
            {allExpanded ? "Свернуть все" : "Развернуть все"}
          </Button>
        )}
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
                <th style={{ ...cellPad, width: 36 }} />
                <th style={cellPad}>КП / группа</th>
                <th style={cellPad}>Заказчик</th>
                <th style={cellPad}>Срок</th>
                <th style={cellPad}>Поз.</th>
                <th style={cellPad}>Σ Qty</th>
                <th style={cellPad}>N/M</th>
                <th style={cellPad} />
              </tr>
            </thead>
            <tbody>
              {groups.map((group) => {
                const isExpanded = expandedGroupKeys.has(group.key);

                return (
                  <Fragment key={String(group.key)}>
                    {renderGroupHeader(group)}
                    {isExpanded && (
                      <tr style={{ background: "#fafbff" }}>
                        <td style={{ padding: 0 }} />
                        <td colSpan={OUTER_COL_COUNT - 1} style={{ padding: "0.5rem 0.75rem 0.85rem" }}>
                          {renderPlateRows(group)}
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
