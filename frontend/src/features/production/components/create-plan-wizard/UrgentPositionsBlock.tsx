import { Fragment, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import type { UrgentPosition } from "@/features/production/types/production";
import { formatRu } from "./utils";
import { subTdStyle, subThStyle, tdStyle, thStyle } from "./tableStyles";

export const urgentPositionKey = (pos: Pick<UrgentPosition, "kp_id" | "plate_id">) =>
  `${pos.kp_id}:${pos.plate_id}`;

const conflictTitle = (conflict: string): string => {
  if (conflict === "schedule_earlier") {
    return "График поставки раньше срока КП более чем на 7 дней";
  }
  if (conflict === "kp_earlier") {
    return "Срок КП раньше графика поставки более чем на 7 дней";
  }
  return conflict;
};

const formatDetailLine = (detail: UrgentPosition["deadline_details"][number]): string => {
  const deadline = detail.deadline ? formatRu(String(detail.deadline)) : "—";
  const qty = typeof detail.qty === "number" ? detail.qty : "—";
  if (detail.type === "delivery_batch") {
    const name = detail.batch_name ? `«${detail.batch_name}»` : "без названия";
    return `Партия ${name}: ${deadline}, ${qty} шт`;
  }
  if (detail.type === "execution_terms") {
    return `Срок КП: ${deadline}, ${qty} шт`;
  }
  return `${detail.type || "источник"}: ${deadline}, ${qty} шт`;
};

export type UrgentPositionsBlockProps = {
  positions: UrgentPosition[];
  selectedPlatesByKp: Record<number, number[]>;
  loading?: boolean;
  errorMessage?: string | null;
  onTogglePosition: (position: UrgentPosition) => void;
};

export const UrgentPositionsBlock = ({
  positions,
  selectedPlatesByKp,
  loading = false,
  errorMessage = null,
  onTogglePosition,
}: UrgentPositionsBlockProps) => {
  const [expandedKeys, setExpandedKeys] = useState<Set<string>>(() => new Set());

  const toggleExpand = (key: string) => {
    setExpandedKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  };

  const isChecked = (pos: UrgentPosition) =>
    (selectedPlatesByKp[pos.kp_id] ?? []).includes(pos.plate_id);

  return (
    <div style={{ display: "grid", gap: "0.5rem" }}>
      <div style={{ fontWeight: 600, color: "#23366f", fontSize: "0.95rem" }}>
        Срочные по срокам
      </div>

      {loading && (
        <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
          <Spinner /> Анализ срочных позиций…
        </div>
      )}

      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}

      {!loading && !errorMessage && positions.length === 0 && (
        <div style={{ color: "#475467", fontSize: "0.9rem" }}>
          Нет позиций с дедлайном в выбранные дни.
        </div>
      )}

      {positions.length > 0 && (
        <div
          style={{
            border: "1px solid #e4e7ec",
            borderRadius: 14,
            overflow: "hidden",
          }}
        >
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead style={{ background: "#f8fafc" }}>
              <tr>
                <th style={thStyle} />
                <th style={thStyle}>Выбор</th>
                <th style={thStyle}>Наименование</th>
                <th style={thStyle}>Дедлайн</th>
                <th style={thStyle}>Остаток</th>
                <th style={thStyle}>КП №</th>
                <th style={thStyle} />
              </tr>
            </thead>
            <tbody>
              {positions.map((pos) => {
                const key = urgentPositionKey(pos);
                const expanded = expandedKeys.has(key);
                const checked = isChecked(pos);
                return (
                  <Fragment key={key}>
                    <tr style={{ borderTop: "1px solid #e4e7ec" }}>
                      <td style={{ ...tdStyle, width: 36, textAlign: "center" }}>
                        <button
                          type="button"
                          onClick={() => toggleExpand(key)}
                          aria-label={
                            expanded
                              ? "Свернуть детали дедлайна"
                              : "Развернуть детали дедлайна"
                          }
                          aria-expanded={expanded}
                          style={{
                            border: "none",
                            background: "transparent",
                            cursor: "pointer",
                            fontSize: "0.95rem",
                            color: "#475467",
                            padding: "0.1rem 0.25rem",
                            lineHeight: 1,
                          }}
                        >
                          {expanded ? "▾" : "▸"}
                        </button>
                      </td>
                      <td style={tdStyle}>
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => onTogglePosition(pos)}
                          aria-label={`Выбрать ${pos.plate_name}`}
                        />
                      </td>
                      <td style={tdStyle}>{pos.plate_name || "—"}</td>
                      <td style={tdStyle}>{formatRu(pos.deadline)}</td>
                      <td style={tdStyle}>{pos.qty_remaining}</td>
                      <td style={tdStyle}>{pos.kp_id}</td>
                      <td style={tdStyle}>
                        {pos.conflict ? (
                          <span title={conflictTitle(pos.conflict)} aria-label="Конфликт сроков">
                            ⚠️
                          </span>
                        ) : null}
                      </td>
                    </tr>
                    {expanded && (
                      <tr style={{ background: "#fafbff" }}>
                        <td style={{ padding: 0 }} />
                        <td colSpan={6} style={{ padding: "0.5rem 0.75rem 0.85rem" }}>
                          {pos.deadline_details.length === 0 ? (
                            <div style={{ color: "#475467" }}>Нет деталей дедлайна.</div>
                          ) : (
                            <table style={{ width: "100%", borderCollapse: "collapse" }}>
                              <thead>
                                <tr>
                                  <th style={subThStyle}>Источник</th>
                                  <th style={subThStyle}>Детали</th>
                                </tr>
                              </thead>
                              <tbody>
                                {pos.deadline_details.map((detail, idx) => (
                                  <tr
                                    key={`${key}-d-${idx}`}
                                    style={{ borderTop: "1px solid #eef2f6" }}
                                  >
                                    <td style={subTdStyle}>
                                      {detail.type === "delivery_batch"
                                        ? "Партия"
                                        : detail.type === "execution_terms"
                                          ? "Срок КП"
                                          : detail.type || "—"}
                                    </td>
                                    <td style={subTdStyle}>{formatDetailLine(detail)}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          )}
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
    </div>
  );
};
