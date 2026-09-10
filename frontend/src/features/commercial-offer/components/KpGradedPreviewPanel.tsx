import { useMemo, useState } from "react";
import type { CommercialDraftDetails, SimpleKpProductType } from "@/features/commercial-offer/types/commercialOffer";
import { PRODUCT_TYPE_CONFIG } from "@/features/commercial-offer/lib/productTypeConfig";
import { getSimpleProductTypePreview } from "@/features/commercial-offer/lib/productTypePreview";
import { formatOfferNumber, formatOfferSum } from "@/features/commercial-offer/lib/formatOfferNumbers";
import { filterCompositionWarnings } from "@/features/commercial-offer/lib/compositionWarnings";
import type { LineRowHandlers } from "@/features/commercial-offer/lib/lineRowHandlers";
import { LineActionsCell, LineActionsHeader } from "@/features/commercial-offer/components/LineRowActions";
import { LineUndoToast } from "@/features/commercial-offer/components/LineUndoToast";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";

type KpGradedPreviewPanelProps = {
  productType: SimpleKpProductType;
  draft: CommercialDraftDetails;
  normalizedText: string;
  isUpdatingGrades?: boolean;
  onApplyGradeToAll?: (grade: string) => void;
  onLineGradeChange?: (lineIndex: number, grade: string) => void;
  lineRowHandlers?: LineRowHandlers;
};

export const KpGradedPreviewPanel = ({
  productType,
  draft,
  normalizedText,
  isUpdatingGrades = false,
  onApplyGradeToAll,
  onLineGradeChange,
  lineRowHandlers,
}: KpGradedPreviewPanelProps) => {
  const config = PRODUCT_TYPE_CONFIG[productType];
  const labels = config.labels;
  const supportsGrades = config.supportsGrades;
  const preview = getSimpleProductTypePreview(productType);
  const rows = useMemo(() => preview.buildRows(draft), [preview, draft]);
  const warnings = filterCompositionWarnings(draft.metadata.warnings ?? []);
  const validationErrors = draft.wizard_state.validation_errors ?? [];
  const defaultGrade = draft.metadata.default_concrete_grade ?? "B25";
  const [bulkGrade, setBulkGrade] = useState(defaultGrade);

  const normalizedTextChanged =
    normalizedText.trim() !== (draft.metadata.normalized_text ?? "").trim() && normalizedText.trim().length > 0;

  const hasUnpricedRows = rows.some((row) => row.unit_price === null);
  const hasEditableGradeRows = rows.some((row) => !row.sealed);

  return (
    <Card
      title="Состав КП (предпросмотр)"
      subtitle={labels.previewSubtitle}
    >
      <div style={{ display: "grid", gap: "0.75rem" }}>
        {lineRowHandlers?.undoToast ? (
          <LineUndoToast
            message={lineRowHandlers.undoToast.message}
            onUndo={lineRowHandlers.undoToast.onUndo}
          />
        ) : null}
        {warnings.length > 0 && (
          <Alert tone="warning">
            <div style={{ display: "grid", gap: "0.35rem" }}>
              {warnings.map((warning) => (
                <div key={warning}>{warning}</div>
              ))}
            </div>
          </Alert>
        )}

        {validationErrors.length > 0 && (
          <Alert tone="error">
            <div style={{ display: "grid", gap: "0.35rem" }}>
              {validationErrors.map((error) => (
                <div key={error}>{error}</div>
              ))}
            </div>
          </Alert>
        )}

        {normalizedTextChanged && (
          <Alert tone="info">{labels.previewChangedMessage}</Alert>
        )}

        {supportsGrades && onApplyGradeToAll && hasEditableGradeRows && (
          <div
            style={{
              display: "flex",
              gap: "0.75rem",
              flexWrap: "wrap",
              alignItems: "center",
              padding: "0.75rem",
              border: "1px solid #e4e7ec",
              borderRadius: 12,
              background: "#fafafa",
            }}
          >
            <span style={{ color: "#475467" }}>Применить класс ко всем новым:</span>
            <select
              value={bulkGrade}
              onChange={(event) => setBulkGrade(event.target.value)}
              disabled={isUpdatingGrades}
              style={{
                border: "1px solid #d0d5dd",
                borderRadius: 8,
                padding: "0.45rem 0.65rem",
                background: "#ffffff",
              }}
            >
              {preview.gradeCodes.map((code) => (
                <option key={code} value={code}>
                  {preview.formatGradeLabel(code)}
                </option>
              ))}
            </select>
            <Button
              type="button"
              variant="secondary"
              disabled={isUpdatingGrades || !preview.isGradeCode(bulkGrade)}
              onClick={() => onApplyGradeToAll(bulkGrade)}
            >
              {isUpdatingGrades ? "Обновляем..." : "Применить"}
            </Button>
          </div>
        )}

        {rows.length === 0 ? (
          <div style={{ color: "#667085" }}>{labels.previewEmptyMessage}</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table
              style={{
                width: "100%",
                minWidth: supportsGrades ? 640 : 560,
                borderCollapse: "collapse",
                tableLayout: "auto",
                fontSize: "0.92rem",
              }}
            >
              <thead>
                <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                  {(supportsGrades
                    ? ["№", "Марка", "Класс", "Кол-во", "Цена", "Сумма"]
                    : ["№", "Марка", "Кол-во", "Цена", "Сумма"]
                  ).map((column) => (
                    <th key={column} style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #e4e7ec" }}>
                      {column}
                    </th>
                  ))}
                  <LineActionsHeader enabled={Boolean(lineRowHandlers)} />
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => {
                  const isUnpriced = row.unit_price === null;
                  return (
                    <tr
                      key={
                        supportsGrades
                          ? `${row.mark}-${row.concrete_grade}-${index}`
                          : `${row.mark}-${index}`
                      }
                    >
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>{index + 1}</td>
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>{row.mark}</td>
                      {supportsGrades && (
                        <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>
                          {onLineGradeChange && !row.sealed ? (
                            <select
                              value={row.concrete_grade}
                              disabled={isUpdatingGrades}
                              onChange={(event) => onLineGradeChange(index, event.target.value)}
                              style={{
                                border: "1px solid #d0d5dd",
                                borderRadius: 8,
                                padding: "0.35rem 0.5rem",
                                background: isUnpriced ? "#fef3f2" : "#ffffff",
                              }}
                            >
                              {(row.available_grades?.length ? row.available_grades : preview.gradeCodes).map(
                                (code) => (
                                  <option key={code} value={code}>
                                    {preview.formatGradeLabel(code)}
                                  </option>
                                ),
                              )}
                            </select>
                          ) : (
                            preview.formatGradeLabel(row.concrete_grade)
                          )}
                        </td>
                      )}
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>{row.qty}</td>
                      <td
                        style={{
                          padding: "0.55rem 0.65rem",
                          borderBottom: "1px solid #f2f4f7",
                          color: isUnpriced ? "#b42318" : "inherit",
                        }}
                      >
                        {isUnpriced ? "нет в прайсе" : formatOfferNumber(row.unit_price)}
                      </td>
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>
                        {isUnpriced ? "—" : formatOfferSum(row.qty, row.unit_price)}
                      </td>
                      <LineActionsCell
                        handlers={lineRowHandlers}
                        lineId={row.lineId}
                        qty={row.qty}
                        sourceText={row.sourceText ?? ""}
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
            {labels.previewUnpricedMessage}
          </Alert>
        )}

      </div>
    </Card>
  );
};
