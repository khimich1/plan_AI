import { useState } from "react";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";
import { Alert } from "@/shared/ui/Alert";
import { useKpReadinessPositionsQuery } from "@/features/commercial-archive/hooks/useArchiveQueries";
import type {
  KpReadinessStep,
  KpReadinessStepState,
  KpReadinessSummary,
} from "@/features/commercial-archive/types/archive";
import { getErrorMessage } from "@/shared/lib/apiError";

type Props = {
  kpId: number;
  readiness: KpReadinessSummary;
};

export const KpReadinessBlock = ({ kpId, readiness }: Props) => {
  const [expanded, setExpanded] = useState(false);
  const [copyFeedback, setCopyFeedback] = useState<string | null>(null);
  const positionsQuery = useKpReadinessPositionsQuery(kpId, { enabled: expanded });

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(readiness.client_copy_text);
      setCopyFeedback("Текст скопирован в буфер обмена");
      window.setTimeout(() => setCopyFeedback(null), 2500);
    } catch {
      setCopyFeedback("Не удалось скопировать — скопируйте вручную");
    }
  };

  const sgp = readiness.sgp_progress;

  return (
    <section
      style={{
        padding: "1rem",
        border: "1px solid #e4e7ec",
        borderRadius: 14,
        background: "#ffffff",
        display: "grid",
        gap: "0.85rem",
      }}
    >
      <h3 style={{ margin: 0, fontSize: "1rem" }}>Статус производства</h3>

      <ReadinessStepper steps={readiness.steps} />

      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: "1.25rem",
          fontSize: "0.9rem",
          color: "#344054",
        }}
      >
        {readiness.completion_percentage !== null && (
          <div>
            <span style={{ color: "#667085" }}>Производство: </span>
            <strong>{readiness.completion_percentage.toFixed(0)}%</strong>
          </div>
        )}
        {sgp && (
          <div>
            <span style={{ color: "#667085" }}>СГП: </span>
            <strong>
              {sgp.n}/{sgp.m}
            </strong>
          </div>
        )}
      </div>

      {readiness.summary_text && (
        <p style={{ margin: 0, color: "#475467", lineHeight: 1.45 }}>{readiness.summary_text}</p>
      )}

      {readiness.expected_sgp_date_label && (
        <p style={{ margin: 0, color: "#475467", lineHeight: 1.45 }}>
          Ожидаем на СГП к: <strong>{readiness.expected_sgp_date_label}</strong>
        </p>
      )}

      {readiness.release_note && (
        <p style={{ margin: 0, fontSize: "0.85rem", color: "#98a2b3" }}>{readiness.release_note}</p>
      )}

      <div style={{ display: "flex", flexWrap: "wrap", gap: "0.5rem", alignItems: "center" }}>
        <Button variant="ghost" onClick={() => setExpanded((prev) => !prev)}>
          {expanded ? "Свернуть ▲" : "Подробнее ▼"}
        </Button>
        <Button variant="secondary" onClick={handleCopy}>
          Скопировать для клиента
        </Button>
        {copyFeedback && (
          <span style={{ fontSize: "0.85rem", color: "#027a48" }}>{copyFeedback}</span>
        )}
      </div>

      {expanded && (
        <div style={{ marginTop: "0.25rem" }}>
          {positionsQuery.isPending && (
            <div style={{ display: "flex", gap: "0.5rem", alignItems: "center", color: "#667085" }}>
              <Spinner /> Загружаю позиции...
            </div>
          )}
          {positionsQuery.isError && (
            <Alert tone="error">{getErrorMessage(positionsQuery.error)}</Alert>
          )}
          {positionsQuery.data && (
            <div style={{ overflowX: "auto" }}>
              {positionsQuery.data.items.length === 0 ? (
                <div style={{ color: "#667085" }}>Нет данных по позициям.</div>
              ) : (
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "0.9rem" }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Позиция</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Заказ</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>В плане</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>На СГП</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Осталось</th>
                    </tr>
                  </thead>
                  <tbody>
                    {positionsQuery.data.items.map((row, index) => (
                      <tr key={`${row.label}-${index}`} style={{ borderTop: "1px solid #e4e7ec" }}>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{row.label}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{row.ordered}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{row.in_plan}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{row.on_sgp}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{row.remaining}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}
        </div>
      )}
    </section>
  );
};

const ReadinessStepper = ({ steps }: { steps: KpReadinessStep[] }) => (
  <div
    style={{
      display: "flex",
      flexWrap: "wrap",
      gap: "0.35rem",
      alignItems: "flex-start",
    }}
  >
    {steps.map((step, index) => (
      <div key={step.id} style={{ display: "flex", alignItems: "center", gap: "0.35rem" }}>
        <StepChip step={step} />
        {index < steps.length - 1 && (
          <span style={{ color: "#d0d5dd", fontSize: "0.75rem" }}>→</span>
        )}
      </div>
    ))}
  </div>
);

const StepChip = ({ step }: { step: KpReadinessStep }) => {
  const colors = stepColors[step.state];
  return (
    <div
      style={{
        padding: "0.35rem 0.6rem",
        borderRadius: 10,
        border: `1px solid ${colors.border}`,
        background: colors.bg,
        color: colors.text,
        fontSize: "0.8rem",
        lineHeight: 1.2,
        minWidth: 72,
        textAlign: "center",
        opacity: step.state === "disabled" ? 0.55 : 1,
      }}
    >
      <div style={{ fontWeight: 600 }}>{step.label}</div>
      {step.hint && (
        <div style={{ fontSize: "0.72rem", marginTop: "0.15rem", opacity: 0.85 }}>{step.hint}</div>
      )}
    </div>
  );
};

const stepColors: Record<
  KpReadinessStepState,
  { bg: string; border: string; text: string }
> = {
  done: { bg: "#ecfdf3", border: "#abefc6", text: "#027a48" },
  active: { bg: "#eef4ff", border: "#b2ddff", text: "#175cd3" },
  pending: { bg: "#f9fafb", border: "#e4e7ec", text: "#667085" },
  disabled: { bg: "#f2f4f7", border: "#e4e7ec", text: "#98a2b3" },
};
