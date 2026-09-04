import { useEffect, useMemo, useRef, useState } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { Modal } from "@/shared/ui/Modal";
import { Drawer } from "@/shared/ui/Drawer";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import { useMoveToProductionMutation } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { isCapacityRed } from "@/features/factory-capacity/components/FactoryCapacityPanel";
import { PromiseKnobSettings } from "@/features/factory-capacity/components/PromiseKnobSettings";
import { PromiseQuoteBlock } from "@/features/factory-capacity/components/PromiseQuoteBlock";
import { PromiseWeekStrip } from "@/features/factory-capacity/components/PromiseWeekStrip";
import {
  holdCreatedByTitle,
  isoToDdMmYyyy,
  promiseHoldKeys,
  promiseQuoteKeys,
  useCreatePromiseHoldMutation,
  usePromiseHoldQuery,
  usePromiseQuoteQuery,
} from "@/features/factory-capacity/api/promiseQuote";
import { useCapacitySnapshotQuery } from "@/features/factory-capacity/hooks/useCapacitySnapshotQuery";
import { ddMmYyyyToIso } from "@/features/factory-capacity/lib/dates";
import { getErrorMessage } from "@/shared/lib/apiError";
import {
  EXECUTION_TERMS_FIELD_HINT,
  EXECUTION_TERMS_PARSE_ERROR,
  EXECUTION_TERMS_PLACEHOLDER,
  tryNormalizeExecutionTerms,
} from "@/shared/lib/executionTerms";

type Props = {
  open: boolean;
  onClose: () => void;
  kpId: number;
  /** Уже сохранённый в карточке КП срок (обычно ДД.ММ.ГГГГ) — показываем как стартовое значение. */
  initialExecutionTerms?: string | null;
};

function useDebouncedValue<T>(value: T, delayMs: number): T {
  const [debounced, setDebounced] = useState(value);
  useEffect(() => {
    const id = window.setTimeout(() => setDebounced(value), delayMs);
    return () => window.clearTimeout(id);
  }, [value, delayMs]);
  return debounced;
}

export const MoveToProductionDialog = ({
  open,
  onClose,
  kpId,
  initialExecutionTerms,
}: Props) => {
  const [value, setValue] = useState("");
  const [capacityOpen, setCapacityOpen] = useState(false);
  const userEditedRef = useRef(false);
  const queryClient = useQueryClient();
  const mutation = useMoveToProductionMutation();
  const holdMutation = useCreatePromiseHoldMutation();
  const quote = usePromiseQuoteQuery(open ? kpId : null);
  const holdQuery = usePromiseHoldQuery(open ? kpId : null);
  const activeHold = holdQuery.data ?? null;

  useEffect(() => {
    if (!open) {
      userEditedRef.current = false;
      setCapacityOpen(false);
      return;
    }
    mutation.reset();
    holdMutation.reset();
    userEditedRef.current = false;
    setValue("");
    setCapacityOpen(false);
  }, [open, kpId]); // eslint-disable-line react-hooks/exhaustive-deps -- только open/kpId: сброс при открытии/смене КП

  useEffect(() => {
    if (!open || userEditedRef.current) {
      return;
    }
    const fromHold = isoToDdMmYyyy(activeHold?.promised_date);
    if (fromHold) {
      setValue(fromHold);
      return;
    }

    const fromCard = (initialExecutionTerms ?? "").trim();
    if (fromCard) {
      setValue(fromCard);
      return;
    }

    if (quote.isPending) {
      setValue("");
      return;
    }

    const promised = isoToDdMmYyyy(quote.data?.window?.promised_date);
    if (promised) {
      setValue(promised);
      return;
    }

    setValue("");
  }, [
    open,
    kpId,
    initialExecutionTerms,
    quote.isPending,
    quote.data?.window?.promised_date,
    activeHold?.promised_date,
  ]);

  const normalizedTerms = useMemo(() => {
    const trimmed = value.trim();
    if (!trimmed) return null;
    return tryNormalizeExecutionTerms(trimmed);
  }, [value]);

  const targetIso = useMemo(
    () => (normalizedTerms ? ddMmYyyyToIso(normalizedTerms) : null),
    [normalizedTerms],
  );
  const debouncedTarget = useDebouncedValue(targetIso, 300);

  const capacity = useCapacitySnapshotQuery(open ? kpId : null, debouncedTarget);
  const capacityBlocked = isCapacityRed(capacity.data);

  const onSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    const trimmed = value.trim();
    if (!trimmed || tryNormalizeExecutionTerms(trimmed) === null) return;
    if (capacityBlocked) return;

    try {
      await mutation.mutateAsync({ kpId, executionTerms: trimmed });
      await queryClient.invalidateQueries({ queryKey: promiseQuoteKeys.all });
      await queryClient.invalidateQueries({ queryKey: promiseHoldKeys.all });
      onClose();
    } catch {
      // ошибка показывается через mutation.error
    }
  };

  const onHold = async () => {
    if (!quote.data?.window || activeHold) {
      return;
    }
    try {
      await holdMutation.mutateAsync(kpId);
    } catch {
      // ошибка показывается через holdMutation.error
    }
  };

  const localParseError =
    value.trim().length > 0 && tryNormalizeExecutionTerms(value.trim()) === null
      ? EXECUTION_TERMS_PARSE_ERROR
      : null;

  const handleModalClose = () => {
    if (capacityOpen) {
      setCapacityOpen(false);
      return;
    }
    onClose();
  };

  return (
    <>
      <Modal open={open} onClose={handleModalClose} title={`В производство: КП №${kpId}`}>
        <form onSubmit={onSubmit} style={{ display: "grid", gap: "1rem" }}>
          {quote.isPending && (
            <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
              <Spinner /> Считаю срок по корзинам…
            </div>
          )}
          {quote.isError ? (
            <Alert tone="error">{getErrorMessage(quote.error)}</Alert>
          ) : null}
          {quote.data ? (
            <>
              <PromiseQuoteBlock quote={quote.data} />
              <PromiseWeekStrip weeks={quote.data.weeks} quoteWindow={quote.data.window} />
            </>
          ) : null}
          {activeHold ? (
            <div
              data-testid="promise-hold-locked"
              role="status"
              title={holdCreatedByTitle(activeHold.created_by)}
              style={{
                border: "1px solid #bfd4ff",
                background: "#eef4ff",
                color: "#1d4ed8",
                borderRadius: 14,
                padding: "0.9rem 1rem",
                fontSize: "0.9rem",
                fontWeight: 600,
              }}
            >
              Срок закреплён до сегодня
              {activeHold.created_by ? ` · ${activeHold.created_by}` : ""}
            </div>
          ) : null}
          <FieldWrapper label="Срок выполнения" hint={EXECUTION_TERMS_FIELD_HINT}>
            <Input
              type="text"
              value={value}
              onChange={(event) => {
                userEditedRef.current = true;
                setValue(event.target.value);
              }}
              placeholder={EXECUTION_TERMS_PLACEHOLDER}
              autoFocus
            />
          </FieldWrapper>
          {localParseError && <Alert tone="error">{localParseError}</Alert>}
          {capacityBlocked && capacity.data?.hint ? (
            <Alert tone="error">
              {capacity.data.hint}. Увеличьте срок в поле выше и повторите.
            </Alert>
          ) : null}
          {mutation.isError && <Alert tone="error">{getErrorMessage(mutation.error)}</Alert>}
          {holdMutation.isError && <Alert tone="error">{getErrorMessage(holdMutation.error)}</Alert>}

          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", flexWrap: "wrap" }}>
            <Button
              variant="ghost"
              type="button"
              onClick={() => setCapacityOpen(true)}
              disabled={!quote.data}
            >
              Ёмкость
            </Button>
            <Button
              variant="ghost"
              type="button"
              onClick={() => {
                void onHold();
              }}
              disabled={
                holdMutation.isPending ||
                mutation.isPending ||
                Boolean(activeHold) ||
                !quote.data?.window
              }
            >
              {holdMutation.isPending ? "Закрепляю…" : "Закрепить срок"}
            </Button>
            <Button variant="ghost" type="button" onClick={handleModalClose} disabled={mutation.isPending}>
              Отмена
            </Button>
            <Button
              type="submit"
              disabled={
                mutation.isPending ||
                !value.trim() ||
                tryNormalizeExecutionTerms(value.trim()) === null ||
                capacityBlocked
              }
            >
              {mutation.isPending ? "Перевод..." : "Перевести в производство"}
            </Button>
          </div>
        </form>
      </Modal>

      <Drawer
        open={open && capacityOpen}
        onClose={() => setCapacityOpen(false)}
        title="Ёмкость завода"
        side="left"
        width={380}
      >
        {quote.data ? (
          <>
            <PromiseWeekStrip weeks={quote.data.weeks} quoteWindow={quote.data.window} />
            <PromiseKnobSettings currentKnob={quote.data.knob} />
          </>
        ) : null}
      </Drawer>
    </>
  );
};
