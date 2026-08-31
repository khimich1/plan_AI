import { useEffect, useMemo, useState } from "react";
import { Modal } from "@/shared/ui/Modal";
import { Drawer } from "@/shared/ui/Drawer";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";
import { Alert } from "@/shared/ui/Alert";
import {
  DeliveryScheduleEditor,
  validateScheduleEditor,
} from "@/features/delivery-schedule/components/DeliveryScheduleEditor";
import { ImportScheduleDialog } from "@/features/delivery-schedule/components/ImportScheduleDialog";
import {
  useDeliveryScheduleQuery,
  usePutDeliveryScheduleMutation,
} from "@/features/delivery-schedule/hooks/useDeliveryScheduleQueries";
import {
  draftsToPut,
  importBatchesToDrafts,
  viewToDrafts,
  type BatchDraft,
  type OfferPlateForSchedule,
} from "@/features/delivery-schedule/lib/scheduleDraft";
import type {
  ImportDraftResponse,
  UnmatchedRowOut,
} from "@/features/delivery-schedule/types/deliverySchedule";
import {
  FactoryCapacityPanel,
  isCapacityRed,
} from "@/features/factory-capacity/components/FactoryCapacityPanel";
import { useCapacitySnapshotQuery } from "@/features/factory-capacity/hooks/useCapacitySnapshotQuery";
import { maxIsoDate } from "@/features/factory-capacity/lib/dates";
import { getErrorMessage } from "@/shared/lib/apiError";

type Props = {
  open: boolean;
  onClose: () => void;
  kpId: number;
  plates: OfferPlateForSchedule[];
  readOnly?: boolean;
};

export const DeliveryScheduleDialog = ({
  open,
  onClose,
  kpId,
  plates,
  readOnly = false,
}: Props) => {
  const query = useDeliveryScheduleQuery(open ? kpId : null);
  const putMutation = usePutDeliveryScheduleMutation();

  const [batches, setBatches] = useState<BatchDraft[]>([]);
  const [invoiceNumber, setInvoiceNumber] = useState("");
  const [contractNumber, setContractNumber] = useState("");
  const [localError, setLocalError] = useState<string | null>(null);
  const [hydratedFor, setHydratedFor] = useState<string | null>(null);
  const [importOpen, setImportOpen] = useState(false);
  const [capacityOpen, setCapacityOpen] = useState(false);
  const [unmatchedRows, setUnmatchedRows] = useState<UnmatchedRowOut[]>([]);

  const hydrateKey = useMemo(() => {
    if (!open) {
      return null;
    }
    if (query.isPending) {
      return null;
    }
    const updated = query.data?.updated_at ?? "empty";
    return `${kpId}:${updated}`;
  }, [open, kpId, query.isPending, query.data?.updated_at]);

  useEffect(() => {
    if (!open) {
      setHydratedFor(null);
      setLocalError(null);
      setImportOpen(false);
      setCapacityOpen(false);
      setUnmatchedRows([]);
      putMutation.reset();
      return;
    }
    if (!hydrateKey || hydrateKey === hydratedFor) {
      return;
    }
    setBatches(viewToDrafts(query.data));
    setInvoiceNumber(query.data?.invoice_number ?? "");
    setContractNumber(query.data?.contract_number ?? "");
    setLocalError(null);
    setUnmatchedRows([]);
    setHydratedFor(hydrateKey);
  }, [open, hydrateKey, hydratedFor, query.data]); // eslint-disable-line react-hooks/exhaustive-deps -- putMutation.reset only on close

  const targetIso = useMemo(
    () => maxIsoDate(batches.map((b) => b.produce_by)),
    [batches],
  );
  const capacity = useCapacitySnapshotQuery(open ? kpId : null, targetIso);
  const anyBatchRed = batches.some((b) => b.status === "red");
  const capacityBlocked = isCapacityRed(capacity.data) || anyBatchRed;

  const handleImported = (result: ImportDraftResponse) => {
    setBatches(importBatchesToDrafts(result.batches));
    setUnmatchedRows(result.unmatched_rows ?? []);
    setLocalError(null);
  };

  const handleSave = async () => {
    const error = validateScheduleEditor(plates, batches);
    if (error) {
      setLocalError(error);
      return;
    }
    if (capacityBlocked) {
      setLocalError(
        capacity.data?.hint ??
          "Завод перегружен до выбранного срока. Сдвиньте produce_by партий.",
      );
      return;
    }
    setLocalError(null);
    try {
      const view = await putMutation.mutateAsync({
        kpId,
        payload: draftsToPut(batches, {
          invoice_number: invoiceNumber.trim() || null,
          contract_number: contractNumber.trim() || null,
        }),
      });
      setBatches(viewToDrafts(view));
      setInvoiceNumber(view.invoice_number ?? "");
      setContractNumber(view.contract_number ?? "");
      setHydratedFor(`${kpId}:${view.updated_at}`);
      setUnmatchedRows([]);
    } catch {
      // mutation.error shown below
    }
  };

  const handleMainClose = () => {
    if (importOpen) {
      return;
    }
    if (capacityOpen) {
      setCapacityOpen(false);
      return;
    }
    onClose();
  };

  const capacityHint =
    capacity.data?.hint ??
    (anyBatchRed
      ? "Есть партия с красным сроком. Сдвиньте produce_by и повторите."
      : null);

  return (
    <>
      <Modal open={open} onClose={handleMainClose} title={`График поставки · КП №${kpId}`} maxWidth={960}>
        {query.isPending && (
          <div style={{ display: "flex", gap: "0.75rem", alignItems: "center" }}>
            <Spinner /> Загружаю график...
          </div>
        )}

        {query.isError && <Alert tone="error">{getErrorMessage(query.error)}</Alert>}

        {!query.isPending && !query.isError && (
          <div style={{ display: "grid", gap: "1rem" }}>
            {query.data === null && (
              <Alert tone="info">Графика ещё нет — создайте партии и сохраните.</Alert>
            )}

            <DeliveryScheduleEditor
              plates={plates}
              batches={batches}
              onBatchesChange={setBatches}
              invoiceNumber={invoiceNumber}
              contractNumber={contractNumber}
              onInvoiceNumberChange={setInvoiceNumber}
              onContractNumberChange={setContractNumber}
              readOnly={readOnly}
              validationError={localError}
              kpId={kpId}
              hasSavedSchedule={query.data != null}
              unmatchedRows={unmatchedRows}
              onDismissUnmatched={() => setUnmatchedRows([])}
              onImportClick={readOnly ? undefined : () => setImportOpen(true)}
              trafficLightDegraded={Boolean(query.data?.traffic_light_degraded)}
            />

            {capacityBlocked && capacityHint ? (
              <Alert tone="error">{capacityHint}</Alert>
            ) : null}

            {putMutation.isError && (
              <Alert tone="error">{getErrorMessage(putMutation.error)}</Alert>
            )}

            <div style={{ display: "flex", gap: "0.5rem", justifyContent: "flex-end", flexWrap: "wrap" }}>
              <Button
                type="button"
                variant="ghost"
                onClick={() => setCapacityOpen(true)}
                disabled={!targetIso && !capacity.data}
              >
                Ёмкость
              </Button>
              <Button type="button" variant="ghost" onClick={handleMainClose}>
                Закрыть
              </Button>
              {!readOnly && (
                <Button
                  type="button"
                  onClick={() => void handleSave()}
                  disabled={putMutation.isPending || capacityBlocked}
                >
                  {putMutation.isPending ? "Сохраняю…" : "Сохранить"}
                </Button>
              )}
            </div>
          </div>
        )}
      </Modal>

      <Drawer
        open={open && capacityOpen}
        onClose={() => setCapacityOpen(false)}
        title="Ёмкость завода"
        side="left"
        width={380}
      >
        <FactoryCapacityPanel
          snapshot={capacity.data}
          isLoading={Boolean(targetIso) && capacity.isFetching && !capacity.data}
          errorMessage={capacity.isError ? getErrorMessage(capacity.error) : null}
        />
      </Drawer>

      {!readOnly && (
        <ImportScheduleDialog
          open={importOpen}
          onClose={() => setImportOpen(false)}
          kpId={kpId}
          onImported={handleImported}
        />
      )}
    </>
  );
};
