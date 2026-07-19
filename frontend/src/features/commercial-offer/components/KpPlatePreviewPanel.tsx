import type { CommercialDraftDetails } from "@/features/commercial-offer/types/commercialOffer";
import { buildKpPreviewRows } from "@/features/commercial-offer/lib/buildKpPreviewRows";
import { formatOfferNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";
import { Alert } from "@/shared/ui/Alert";
import { Card } from "@/shared/ui/Card";

type KpPlatePreviewPanelProps = {
  draft: CommercialDraftDetails;
  normalizedText: string;
};

const flagLabel = (flag: "wide_direct" | "wide_split"): string =>
  flag === "wide_direct" ? "Шире стандартной" : "Разделена на стандартные позиции";

export const KpPlatePreviewPanel = ({ draft, normalizedText }: KpPlatePreviewPanelProps) => {
  const rows = buildKpPreviewRows(draft);
  const wideLines = draft.metadata.wide_plate_lines ?? [];
  const unparsedLines = draft.metadata.unparsed_lines ?? [];
  const warnings = draft.metadata.warnings ?? [];
  const showWideAlert = wideLines.length > 0 && !draft.metadata.wide_plates_resolved;
  const normalizedTextChanged =
    normalizedText.trim() !== (draft.metadata.normalized_text ?? "").trim() && normalizedText.trim().length > 0;

  return (
    <Card
      title="Состав КП (предпросмотр)"
      subtitle="Наименование, количество и цена — как в документе. Скидка и доставка учитываются позже."
    >
      <div style={{ display: "grid", gap: "0.75rem" }}>
        {showWideAlert && (
          <Alert tone="warning">
            {wideLines.length}{" "}
            {wideLines.length === 1 ? "позиция шире стандартной" : "позиций шире стандартной"} в списке —
            примите решение в блоке ниже.
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
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => (
                  <tr key={`${row.name}-${index}`}>
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
                      }}
                    >
                      {formatOfferNumber(row.unitPrice)}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
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
