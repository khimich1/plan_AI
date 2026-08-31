import { useState } from "react";
import { Button } from "@/shared/ui/Button";
import { Alert } from "@/shared/ui/Alert";
import {
  useDownloadDeliveryScheduleDocumentMutation,
  useDownloadDeliveryScheduleTemplateMutation,
} from "@/features/delivery-schedule/hooks/useDeliveryScheduleQueries";
import {
  TEMPLATE_FILENAME,
  buildScheduleTemplateXlsx,
} from "@/features/delivery-schedule/lib/buildScheduleTemplateXlsx";
import type { BatchDraft, OfferPlateForSchedule } from "@/features/delivery-schedule/lib/scheduleDraft";
import { getErrorMessage } from "@/shared/lib/apiError";
import { saveBlobAs } from "@/shared/lib/downloadFile";

type Props = {
  kpId: number;
  /** XLSX/PDF доступны только после сохранения графика. */
  documentsDisabled?: boolean;
  compact?: boolean;
  /** Черновик открытого редактора: шаблон собирается на клиенте, без GET /template. */
  plates?: OfferPlateForSchedule[];
  batches?: BatchDraft[];
};

const compactStyle = {
  padding: "0.45rem 0.75rem",
  fontSize: "0.85rem",
  borderRadius: 10,
} as const;

export const ScheduleDocumentButtons = ({
  kpId,
  documentsDisabled = false,
  compact = false,
  plates,
  batches,
}: Props) => {
  const templateMutation = useDownloadDeliveryScheduleTemplateMutation();
  const documentMutation = useDownloadDeliveryScheduleDocumentMutation();
  const [error, setError] = useState<string | null>(null);
  const [clientTemplatePending, setClientTemplatePending] = useState(false);

  const useClientTemplate = plates != null && batches != null;
  const templatePending = useClientTemplate
    ? clientTemplatePending
    : templateMutation.isPending;
  const busy =
    templatePending ||
    (documentMutation.isPending && documentMutation.variables?.kpId === kpId);

  const btnStyle = compact ? compactStyle : undefined;

  const run = async (action: () => Promise<unknown>) => {
    setError(null);
    try {
      await action();
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const downloadTemplate = async () => {
    if (useClientTemplate) {
      setClientTemplatePending(true);
      try {
        const blob = await buildScheduleTemplateXlsx(plates, batches);
        saveBlobAs(blob, TEMPLATE_FILENAME);
      } finally {
        setClientTemplatePending(false);
      }
      return;
    }
    await templateMutation.mutateAsync(kpId);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "0.5rem", alignItems: "flex-start" }}>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "0.45rem" }}>
        <Button
          type="button"
          variant="secondary"
          style={btnStyle}
          disabled={busy}
          onClick={() => void run(downloadTemplate)}
        >
          {templatePending ? "Шаблон…" : "Скачать шаблон"}
        </Button>
        <Button
          type="button"
          variant="secondary"
          style={btnStyle}
          disabled={busy || documentsDisabled}
          title={documentsDisabled ? "Сначала сохраните график" : undefined}
          onClick={() =>
            void run(() => documentMutation.mutateAsync({ kpId, fmt: "xlsx" }))
          }
        >
          {documentMutation.isPending && documentMutation.variables?.fmt === "xlsx"
            ? "XLSX…"
            : "XLSX"}
        </Button>
        <Button
          type="button"
          variant="secondary"
          style={btnStyle}
          disabled={busy || documentsDisabled}
          title={documentsDisabled ? "Сначала сохраните график" : undefined}
          onClick={() =>
            void run(() => documentMutation.mutateAsync({ kpId, fmt: "pdf" }))
          }
        >
          {documentMutation.isPending && documentMutation.variables?.fmt === "pdf"
            ? "PDF…"
            : "PDF"}
        </Button>
      </div>
      {error && <Alert tone="error">{error}</Alert>}
    </div>
  );
};
