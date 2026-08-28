import { useEffect, useState, type CSSProperties, type FormEvent } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { useDownloadGsmUsageReportMutation } from "@/features/gsm/hooks/useGsmQueries";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";

const formStyle: CSSProperties = { display: "grid", gap: "0.85rem" };

const labelStyle: CSSProperties = {
  display: "grid",
  gap: 4,
  fontSize: "0.85rem",
  color: "#475467",
};

const actionsStyle: CSSProperties = {
  display: "flex",
  gap: "0.75rem",
  justifyContent: "flex-end",
  flexWrap: "wrap",
};

type Props = {
  open: boolean;
  periodFrom: string;
  periodTo: string;
  vehicleIds: number[] | null;
  onClose: () => void;
  onDownloaded?: () => void;
};

export const UsageReportDialog = ({
  open,
  periodFrom,
  periodTo,
  vehicleIds,
  onClose,
  onDownloaded,
}: Props) => {
  const [from, setFrom] = useState(periodFrom);
  const [to, setTo] = useState(periodTo);
  const [error, setError] = useState<string | null>(null);
  const downloadMutation = useDownloadGsmUsageReportMutation();

  useEffect(() => {
    if (!open) {
      return;
    }
    setFrom(periodFrom);
    setTo(periodTo);
    setError(null);
  }, [open, periodFrom, periodTo]);

  const handleSubmit = async (event: FormEvent) => {
    event.preventDefault();
    setError(null);
    try {
      await downloadMutation.mutateAsync({
        period_from: from,
        period_to: to,
        vehicle_ids: vehicleIds,
      });
      onDownloaded?.();
      onClose();
    } catch (err) {
      setError(formatGsmError(err));
    }
  };

  const selectedHint =
    vehicleIds == null || vehicleIds.length === 0
      ? "Все активные машины"
      : `Выбрано машин: ${vehicleIds.length}`;

  return (
    <Modal open={open} onClose={onClose} title="Отчёт об использовании ГСМ" maxWidth={440}>
      <form style={formStyle} onSubmit={(e) => void handleSubmit(e)}>
        <p style={{ margin: 0, color: "#475467", fontSize: "0.9rem" }}>
          Скачать zip с отчётами и путевыми листами за период.
        </p>
        <label style={labelStyle}>
          Период с
          <Input
            type="date"
            aria-label="Период с"
            value={from}
            onChange={(e) => setFrom(e.target.value)}
            required
          />
        </label>
        <label style={labelStyle}>
          Период по
          <Input
            type="date"
            aria-label="Период по"
            value={to}
            onChange={(e) => setTo(e.target.value)}
            required
          />
        </label>
        <p style={{ margin: 0, color: "#667085", fontSize: "0.85rem" }}>{selectedHint}</p>
        {error && <Alert tone="error">{error}</Alert>}
        <div style={actionsStyle}>
          <Button type="button" variant="secondary" onClick={onClose} disabled={downloadMutation.isPending}>
            Отмена
          </Button>
          <Button type="submit" disabled={downloadMutation.isPending || !from || !to}>
            {downloadMutation.isPending ? "Скачивание…" : "Скачать отчёт"}
          </Button>
        </div>
      </form>
    </Modal>
  );
};
