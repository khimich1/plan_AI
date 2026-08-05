import { useMemo, useState } from "react";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import {
  buildBridgePileLinesFromOrderData,
  buildBridgePilePreviewRows,
} from "@/features/commercial-offer/lib/buildBridgePilePreviewRows";
import { formatOfferNumber, formatOfferSum } from "@/features/commercial-offer/lib/formatOfferNumbers";
import {
  formatBridgePileGradeLabel,
  isBridgePileGradeCode,
  BRIDGE_PILE_GRADE_CODES,
} from "@/features/commercial-offer/lib/bridgePileGrades";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";

type KpBridgePilePreviewPanelProps = {
  draft: CommercialDraftDetails;
  normalizedText: string;
  isUpdatingGrades?: boolean;
  onApplyGradeToAll?: (grade: string) => void;
  onLineGradeChange?: (lineIndex: number, grade: string) => void;
};

export const KpBridgePilePreviewPanel = ({
  draft,
  normalizedText,
  isUpdatingGrades = false,
  onApplyGradeToAll,
  onLineGradeChange,
}: KpBridgePilePreviewPanelProps) => {
  const rows = useMemo(() => buildBridgePilePreviewRows(draft), [draft]);
  const unparsedLines = draft.metadata.unparsed_lines ?? [];
  const warnings = draft.metadata.warnings ?? [];
  const validationErrors = draft.wizard_state.validation_errors ?? [];
  const defaultGrade = draft.metadata.default_concrete_grade ?? "B25";
  const [bulkGrade, setBulkGrade] = useState(defaultGrade);

  const normalizedTextChanged =
    normalizedText.trim() !== (draft.metadata.normalized_text ?? "").trim() && normalizedText.trim().length > 0;

  const hasUnpricedRows = rows.some((row) => row.unit_price === null);

  return (
    <Card
      title="Состав КП (предпросмотр)"
      subtitle="Марка, класс бетона, количество и цена — как в документе."
    >
      <div style={{ display: "grid", gap: "0.75rem" }}>
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
          <Alert tone="info">Изменён список мостовых свай — нажмите «Список верен» для пересчёта состава.</Alert>
        )}

        {onApplyGradeToAll && rows.length > 0 && (
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
            <span style={{ color: "#475467" }}>Применить класс ко всем:</span>
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
              {BRIDGE_PILE_GRADE_CODES.map((code) => (
                <option key={code} value={code}>
                  {formatBridgePileGradeLabel(code)}
                </option>
              ))}
            </select>
            <Button
              type="button"
              variant="secondary"
              disabled={isUpdatingGrades || !isBridgePileGradeCode(bulkGrade)}
              onClick={() => onApplyGradeToAll(bulkGrade)}
            >
              {isUpdatingGrades ? "Обновляем..." : "Применить"}
            </Button>
          </div>
        )}

        {rows.length === 0 ? (
          <div style={{ color: "#667085" }}>Список пуст — распознайте мостовые сваи.</div>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table
              style={{
                width: "100%",
                minWidth: 640,
                borderCollapse: "collapse",
                tableLayout: "auto",
                fontSize: "0.92rem",
              }}
            >
              <thead>
                <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                  {["№", "Марка", "Класс", "Кол-во", "Цена", "Сумма"].map((column) => (
                    <th key={column} style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #e4e7ec" }}>
                      {column}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => {
                  const isUnpriced = row.unit_price === null;
                  return (
                    <tr key={`${row.mark}-${row.concrete_grade}-${index}`}>
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>{index + 1}</td>
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>{row.mark}</td>
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>
                        {onLineGradeChange ? (
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
                            {(row.available_grades?.length
                              ? row.available_grades
                              : BRIDGE_PILE_GRADE_CODES
                            ).map((code) => (
                              <option key={code} value={code}>
                                {formatBridgePileGradeLabel(code)}
                              </option>
                            ))}
                          </select>
                        ) : (
                          formatBridgePileGradeLabel(row.concrete_grade)
                        )}
                      </td>
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
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {hasUnpricedRows && (
          <Alert tone="error">
            Не все марки найдены в прайсе — исправьте список или класс бетона перед переходом к клиенту.
          </Alert>
        )}

        {unparsedLines.length > 0 && (
          <div style={{ borderTop: "1px solid #e4e7ec", paddingTop: "0.75rem" }}>
            <div style={{ fontWeight: 600, marginBottom: "0.35rem" }}>Не попали в состав</div>
            <ul style={{ margin: 0, paddingLeft: "1.25rem", color: "#667085", fontSize: "0.9rem" }}>
              {unparsedLines.map((line) => (
                <li key={line}>{line}</li>
              ))}
            </ul>
          </div>
        )}
      </div>
    </Card>
  );
};

export { buildBridgePileLinesFromOrderData };
