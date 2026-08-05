import { useMemo } from "react";
import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import { buildStepPreviewRows } from "@/features/commercial-offer/lib/buildStepPreviewRows";
import { formatOfferNumber, formatOfferSum } from "@/features/commercial-offer/lib/formatOfferNumbers";
import { Alert } from "@/shared/ui/Alert";
import { Card } from "@/shared/ui/Card";

type KpStepPreviewPanelProps = {
  draft: CommercialDraftDetails;
  normalizedText: string;
};

export const KpStepPreviewPanel = ({ draft, normalizedText }: KpStepPreviewPanelProps) => {
  const rows = useMemo(() => buildStepPreviewRows(draft), [draft]);
  const unparsedLines = draft.metadata.unparsed_lines ?? [];
  const warnings = draft.metadata.warnings ?? [];
  const validationErrors = draft.wizard_state.validation_errors ?? [];

  const normalizedTextChanged =
    normalizedText.trim() !== (draft.metadata.normalized_text ?? "").trim() && normalizedText.trim().length > 0;

  const hasUnpricedRows = rows.some((row) => row.unit_price === null);

  return (
    <Card
      title="Состав КП (предпросмотр)"
      subtitle="Марка, количество и цена — как в документе."
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
          <Alert tone="info">Изменён список ступеней — нажмите «Список верен» для пересчёта состава.</Alert>
        )}

        {rows.length === 0 ? (
          <div style={{ color: "#667085" }}>Список пуст — распознайте ступени.</div>
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
                  {["№", "Марка", "Кол-во", "Цена", "Сумма"].map((column) => (
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
                    <tr key={`${row.mark}-${index}`}>
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>{index + 1}</td>
                      <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7" }}>{row.mark}</td>
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
            Не все марки найдены в прайсе — исправьте список перед переходом к клиенту.
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
