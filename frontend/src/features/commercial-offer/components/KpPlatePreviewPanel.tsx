import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import { buildKpPreviewRows } from "@/features/commercial-offer/lib/buildKpPreviewRows";
import { filterCompositionWarnings } from "@/features/commercial-offer/lib/compositionWarnings";
import { formatOfferNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";
import type { LineRowHandlers } from "@/features/commercial-offer/lib/lineRowHandlers";
import { LineActionsCell, LineActionsHeader } from "@/features/commercial-offer/components/LineRowActions";
import { LineUndoToast } from "@/features/commercial-offer/components/LineUndoToast";
import { Alert } from "@/shared/ui/Alert";
import { Card } from "@/shared/ui/Card";

type KpPlatePreviewPanelProps = {
  draft: CommercialDraftDetails;
  normalizedText: string;
  lineRowHandlers?: LineRowHandlers;
};

const flagLabel = (flag: "wide_direct" | "wide_split"): string =>
  flag === "wide_direct" ? "Шире стандартной" : "Разделена на стандартные позиции";

export const KpPlatePreviewPanel = ({ draft, normalizedText, lineRowHandlers }: KpPlatePreviewPanelProps) => {
  const rows = buildKpPreviewRows(draft);
  const wideLines = draft.metadata.wide_plate_lines ?? [];
  const invalidWidthLines = draft.metadata.invalid_width_lines ?? [];
  const warnings = filterCompositionWarnings(draft.metadata.warnings ?? []);
  const showWideAlert = wideLines.length > 0 && !draft.metadata.wide_plates_resolved;
  const showInvalidWidthAlert = invalidWidthLines.length > 0 && !draft.metadata.invalid_widths_resolved;
  const invalidWidthKeys = new Set(
    invalidWidthLines.flatMap((item) =>
      [item.line, item.name]
        .filter(Boolean)
        .map((value) => value.trim().toLowerCase().replace(/\s+\d+$/, "")),
    ),
  );
  const isInvalidWidthRow = (name: string) => {
    const key = name.trim().toLowerCase().replace(/\s+\d+$/, "");
    if (invalidWidthKeys.has(key)) {
      return true;
    }
    for (const invalidKey of invalidWidthKeys) {
      if (key.includes(invalidKey) || invalidKey.includes(key)) {
        return true;
      }
    }
    return false;
  };
  const hasUnpricedRows = rows.some((row) => row.unitPrice === null);
  const normalizedTextChanged =
    normalizedText.trim() !== (draft.metadata.normalized_text ?? "").trim() && normalizedText.trim().length > 0;

  return (
    <Card
      title="Состав КП (предпросмотр)"
      subtitle="Наименование, количество и цена — как в документе. Скидка и доставка учитываются позже."
    >
      <div style={{ display: "grid", gap: "0.75rem" }}>
        {lineRowHandlers?.undoToast ? (
          <LineUndoToast
            message={lineRowHandlers.undoToast.message}
            onUndo={lineRowHandlers.undoToast.onUndo}
          />
        ) : null}
        {showWideAlert && (
          <Alert tone="warning">
            {wideLines.length}{" "}
            {wideLines.length === 1 ? "позиция шире стандартной" : "позиций шире стандартной"} в списке —
            примите решение в блоке ниже.
          </Alert>
        )}
        {showInvalidWidthAlert && (
          <Alert tone="warning">
            {invalidWidthLines.length}{" "}
            {invalidWidthLines.length === 1
              ? "позиция с шириной вне таблицы резов"
              : "позиций с шириной вне таблицы резов"}{" "}
            — замените на заводской рез или исключите ниже.
          </Alert>
        )}

        {warnings.length > 0 && (
          <Alert tone="warning">
            <div style={{ display: "grid", gap: "0.35rem" }}>
              {warnings.map((warning) => (
                <div key={warning}>{warning}</div>
              ))}
            </div>
          </Alert>
        )}

        {normalizedTextChanged && (
          <Alert tone="info">Изменён список плит — нажмите «Список верен» для пересчёта состава.</Alert>
        )}

        {rows.length === 0 ? (
          <div style={{ color: "#667085" }}>Список пуст — распознайте плиты.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table
              style={{
                width: "100%",
                minWidth: 560,
                borderCollapse: "collapse",
                tableLayout: "auto",
                fontSize: "0.92rem",
              }}
            >
              <thead>
                <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                  <th
                    style={{
                      padding: "0.55rem 0.65rem",
                      borderBottom: "1px solid #e4e7ec",
                      width: "1%",
                      whiteSpace: "nowrap",
                    }}
                  >
                    №
                  </th>
                  <th style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #e4e7ec", minWidth: 280 }}>
                    Наименование
                  </th>
                  <th
                    style={{
                      padding: "0.55rem 0.65rem",
                      borderBottom: "1px solid #e4e7ec",
                      width: "1%",
                      whiteSpace: "nowrap",
                    }}
                  >
                    Кол-во
                  </th>
                  <th
                    style={{
                      padding: "0.55rem 0.65rem",
                      borderBottom: "1px solid #e4e7ec",
                      textAlign: "right",
                      whiteSpace: "nowrap",
                    }}
                  >
                    Цена
                  </th>
                  {lineRowHandlers ? <LineActionsHeader enabled /> : null}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => {
                  const isUnpriced = row.unitPrice === null;
                  const isInvalidWidth =
                    showInvalidWidthAlert && !row.flag && isInvalidWidthRow(row.name);
                  return (
                    <tr
                      key={`${row.name}-${index}`}
                      style={isInvalidWidth ? { background: "#fef3f2" } : undefined}
                    >
                      <td
                        style={{
                          padding: "0.55rem 0.65rem",
                          borderBottom: "1px solid #f2f4f7",
                          fontVariantNumeric: "tabular-nums",
                          whiteSpace: "nowrap",
                          width: "1%",
                          verticalAlign: "top",
                        }}
                      >
                        {index + 1}
                      </td>
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7", verticalAlign: "top" }}>
                        <div style={{ whiteSpace: "nowrap" }}>{row.name}</div>
                        {row.flag && (
                          <div style={{ marginTop: "0.25rem", color: "#b54708", fontSize: "0.82rem" }}>
                            ⚠ {flagLabel(row.flag)}
                            {row.sourceLine ? ` · из «${row.sourceLine}»` : ""}
                          </div>
                        )}
                        {isInvalidWidth && (
                          <div style={{ marginTop: "0.25rem", color: "#b42318", fontSize: "0.82rem" }}>
                            Нестандартная ширина — решение ниже
                          </div>
                        )}
                      </td>
                      <td
                        style={{
                          padding: "0.55rem 0.65rem",
                          borderBottom: "1px solid #f2f4f7",
                          fontVariantNumeric: "tabular-nums",
                          whiteSpace: "nowrap",
                          width: "1%",
                        }}
                      >
                        {row.qty}
                      </td>
                      <td
                        style={{
                          padding: "0.55rem 0.65rem",
                          paddingRight: "0.75rem",
                          borderBottom: "1px solid #f2f4f7",
                          fontVariantNumeric: "tabular-nums",
                          whiteSpace: "nowrap",
                          textAlign: "right",
                          color: isUnpriced ? "#b42318" : "inherit",
                        }}
                      >
                        {isUnpriced ? "нет в прайсе" : formatOfferNumber(row.unitPrice)}
                      </td>
                      <LineActionsCell
                        handlers={lineRowHandlers}
                        lineId={row.lineId}
                        qty={row.qty}
                        sourceText={row.sourceText}
                      />
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {hasUnpricedRows && (
          <Alert tone="error">
            Не все плиты найдены в прайсе — исправьте список перед переходом к клиенту.
          </Alert>
        )}
      </div>
    </Card>
  );
};
