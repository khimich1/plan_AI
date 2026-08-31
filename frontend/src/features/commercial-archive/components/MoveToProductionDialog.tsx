import { useEffect, useMemo, useRef, useState } from "react";
import { Modal } from "@/shared/ui/Modal";
import { Drawer } from "@/shared/ui/Drawer";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Alert } from "@/shared/ui/Alert";
import { ProductionEstimateAlert } from "@/shared/ui/ProductionEstimateAlert";
import { Spinner } from "@/shared/ui/Spinner";
import {
  useMoveToProductionMutation,
  useProductionEstimateQuery,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import {
  FactoryCapacityPanel,
  isCapacityRed,
} from "@/features/factory-capacity/components/FactoryCapacityPanel";
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
  const mutation = useMoveToProductionMutation();
  const estimate = useProductionEstimateQuery(open ? kpId : null);

  useEffect(() => {
    if (!open) {
      userEditedRef.current = false;
      setCapacityOpen(false);
      return;
    }
    mutation.reset();
    userEditedRef.current = false;
    setValue("");
    setCapacityOpen(false);
  }, [open, kpId]); // eslint-disable-line react-hooks/exhaustive-deps -- только open/kpId: сброс при открытии/смене КП

  useEffect(() => {
    if (!open || userEditedRef.current) {
      return;
    }
    const fromCard = (initialExecutionTerms ?? "").trim();
    if (fromCard) {
      setValue(fromCard);
      return;
    }

    if (estimate.isPending) {
      setValue("");
      return;
    }

    const days = estimate.data?.estimated_days;
    if (days != null && Number.isFinite(days)) {
      setValue(`${days} дней`);
      return;
    }

    setValue("");
  }, [
    open,
    kpId,
    initialExecutionTerms,
    estimate.isPending,
    estimate.data?.estimated_days,
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
      onClose();
    } catch {
      // ошибка показывается через mutation.error
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
          {estimate.isPending && (
            <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
              <Spinner /> Подсчитываю оценку производства...
            </div>
          )}
          {estimate.data && (
            <ProductionEstimateAlert
              estimatedTracks={estimate.data.estimated_tracks}
              estimatedDays={estimate.data.estimated_days}
              totalLengthM={estimate.data.total_length_m}
            />
          )}
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

          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", flexWrap: "wrap" }}>
            <Button
              variant="ghost"
              type="button"
              onClick={() => setCapacityOpen(true)}
              disabled={!debouncedTarget && !capacity.data}
            >
              Ёмкость
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
        <FactoryCapacityPanel
          snapshot={capacity.data}
          isLoading={Boolean(debouncedTarget) && capacity.isFetching && !capacity.data}
          errorMessage={capacity.isError ? getErrorMessage(capacity.error) : null}
        />
      </Drawer>
    </>
  );
};
