import { useState } from "react";
import { Button } from "@/shared/ui/Button";
import { getErrorMessage } from "@/shared/lib/apiError";
import { Alert } from "@/shared/ui/Alert";
import {
  usePatchSupplyContractMutation,
  useSupplyContractRegistryQuery,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import {
  CANCELLED_CONTRACT_STATUS,
  SUPPLY_CONTRACT_SCAN_NOTES,
  SUPPLY_CONTRACT_STATUSES,
  type SupplyContractRegistryRow,
} from "@/features/commercial-archive/types/supplyContract";

type Props = {
  open: boolean;
  onRequestNew: (row: SupplyContractRegistryRow) => void;
};

const formatContractDate = (iso: string): string => {
  const [year, month, day] = iso.split("-");
  if (!year || !month || !day) {
    return iso;
  }
  return `${day}.${month}.${year}`;
};

export const SupplyContractRegistry = ({ open, onRequestNew }: Props) => {
  const query = useSupplyContractRegistryQuery(open);
  const patch = usePatchSupplyContractMutation();
  const [error, setError] = useState<string | null>(null);
  const rows = query.data ?? [];

  const changeStatus = async (row: SupplyContractRegistryRow, status: string) => {
    setError(null);
    try {
      await patch.mutateAsync({ contractId: row.id, payload: { status } });
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const changeScanNote = async (row: SupplyContractRegistryRow, scanNote: string) => {
    setError(null);
    try {
      await patch.mutateAsync({ contractId: row.id, payload: { scan_note: scanNote } });
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  const attachScan = async (row: SupplyContractRegistryRow, file: File | undefined) => {
    if (!file) {
      return;
    }
    setError(null);
    try {
      await patch.mutateAsync({ contractId: row.id, payload: { file } });
    } catch (err) {
      setError(getErrorMessage(err));
    }
  };

  return (
    <section aria-label="Реестр договоров">
      <h3 style={{ margin: "0 0 0.75rem" }}>Реестр договоров</h3>
      {query.isPending && <div>Загрузка реестра…</div>}
      {query.isError && <Alert tone="error">{getErrorMessage(query.error)}</Alert>}
      {error && <Alert tone="error">{error}</Alert>}
      <div style={{ overflowX: "auto" }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "0.95rem" }}>
          <thead>
            <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
              <th style={{ padding: "0.5rem 0.75rem" }}>Номер</th>
              <th style={{ padding: "0.5rem 0.75rem" }}>Дата</th>
              <th style={{ padding: "0.5rem 0.75rem" }}>Контрагент</th>
              <th style={{ padding: "0.5rem 0.75rem" }}>Менеджер</th>
              <th style={{ padding: "0.5rem 0.75rem" }}>Оригинал (либо ЭДО)</th>
              <th style={{ padding: "0.5rem 0.75rem" }}>Наличие скана</th>
              <th style={{ padding: "0.5rem 0.75rem" }} />
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row.id} style={{ borderTop: "1px solid #e4e7ec" }}>
                <td style={{ padding: "0.5rem 0.75rem" }}>{row.number}</td>
                <td style={{ padding: "0.5rem 0.75rem" }}>{formatContractDate(row.contract_date)}</td>
                <td style={{ padding: "0.5rem 0.75rem" }}>{row.counterparty_name || "—"}</td>
                <td style={{ padding: "0.5rem 0.75rem" }}>{row.manager_name}</td>
                <td style={{ padding: "0.5rem 0.75rem" }}>
                  <select
                    aria-label={`Статус ${row.number}`}
                    value={row.status}
                    onChange={(event) => void changeStatus(row, event.target.value)}
                  >
                    {SUPPLY_CONTRACT_STATUSES.map((status) => (
                      <option key={status} value={status}>
                        {status}
                      </option>
                    ))}
                  </select>
                </td>
                <td style={{ padding: "0.5rem 0.75rem" }}>
                  <select
                    aria-label={`Отметка скана ${row.number}`}
                    value={row.scan_note ?? ""}
                    onChange={(event) => void changeScanNote(row, event.target.value)}
                  >
                    {SUPPLY_CONTRACT_SCAN_NOTES.map((note) => (
                      <option key={note || "empty"} value={note}>
                        {note || "—"}
                      </option>
                    ))}
                  </select>
                  <div style={{ marginTop: "0.35rem", color: "#667085" }}>
                    {row.has_scan ? "файл есть" : "файла нет"}
                  </div>
                  <input
                    aria-label={`Файл скана ${row.number}`}
                    type="file"
                    accept="application/pdf,image/jpeg,image/png"
                    onChange={(event) => void attachScan(row, event.target.files?.[0])}
                  />
                </td>
                <td style={{ padding: "0.5rem 0.75rem" }}>
                  {row.status === CANCELLED_CONTRACT_STATUS && (
                    <Button type="button" variant="secondary" onClick={() => onRequestNew(row)}>
                      Новый договор
                    </Button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
};
