import { useEffect, useRef, useState } from "react";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Alert } from "@/shared/ui/Alert";
import { ProductionEstimateAlert } from "@/shared/ui/ProductionEstimateAlert";
import { Spinner } from "@/shared/ui/Spinner";
import {
  useMoveToProductionMutation,
  useProductionEstimateQuery,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
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

export const MoveToProductionDialog = ({
  open,
  onClose,
  kpId,
  initialExecutionTerms,
}: Props) => {
  const [value, setValue] = useState("");
  const userEditedRef = useRef(false);
  const mutation = useMoveToProductionMutation();
  const estimate = useProductionEstimateQuery(open ? kpId : null);

  useEffect(() => {
    if (!open) {
      userEditedRef.current = false;
      return;
    }
    mutation.reset();
    userEditedRef.current = false;
    setValue("");
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

  const onSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    const trimmed = value.trim();
    if (!trimmed || tryNormalizeExecutionTerms(trimmed) === null) return;

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

  return (
    <Modal open={open} onClose={onClose} title={`В производство: КП №${kpId}`}>
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
        {mutation.isError && <Alert tone="error">{getErrorMessage(mutation.error)}</Alert>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button variant="ghost" type="button" onClick={onClose} disabled={mutation.isPending}>
            Отмена
          </Button>
          <Button
            type="submit"
            disabled={
              mutation.isPending ||
              !value.trim() ||
              tryNormalizeExecutionTerms(value.trim()) === null
            }
          >
            {mutation.isPending ? "Перевод..." : "Перевести в производство"}
          </Button>
        </div>
      </form>
    </Modal>
  );
};
