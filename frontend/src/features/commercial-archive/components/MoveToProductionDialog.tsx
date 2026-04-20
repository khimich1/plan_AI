import { useEffect, useState } from "react";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import {
  useMoveToProductionMutation,
  useProductionEstimateQuery,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import { getErrorMessage } from "@/shared/lib/apiError";

type Props = {
  open: boolean;
  onClose: () => void;
  kpId: number;
};

export const MoveToProductionDialog = ({ open, onClose, kpId }: Props) => {
  const [value, setValue] = useState("");
  const mutation = useMoveToProductionMutation();
  const estimate = useProductionEstimateQuery(open ? kpId : null);

  useEffect(() => {
    if (!open) {
      return;
    }
    mutation.reset();
    if (estimate.data?.estimated_days) {
      setValue(`${estimate.data.estimated_days} дней`);
    } else {
      setValue("");
    }
  }, [open, estimate.data?.estimated_days, mutation]);

  const onSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!value.trim()) {
      return;
    }
    try {
      await mutation.mutateAsync({ kpId, executionTerms: value.trim() });
      onClose();
    } catch {
      // ошибка показывается через mutation.error
    }
  };

  return (
    <Modal open={open} onClose={onClose} title={`В производство: КП №${kpId}`}>
      <form onSubmit={onSubmit} style={{ display: "grid", gap: "1rem" }}>
        {estimate.isPending && (
          <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
            <Spinner /> Подсчитываю оценку производства...
          </div>
        )}
        {estimate.data && (
          <Alert tone="info">
            Оценка производства: ~{estimate.data.estimated_tracks} дорожек, ~{estimate.data.estimated_days} дней
            (суммарная длина {estimate.data.total_length_m.toFixed(1)} м).
          </Alert>
        )}
        <FieldWrapper
          label="Срок выполнения"
          hint="Формат: ДД.ММ.ГГГГ, ГГГГ-ММ-ДД, «7 дней» или «2 недели»"
        >
          <Input
            type="text"
            value={value}
            onChange={(event) => setValue(event.target.value)}
            placeholder="например, 14 дней"
            autoFocus
          />
        </FieldWrapper>
        {mutation.isError && <Alert tone="error">{getErrorMessage(mutation.error)}</Alert>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button variant="ghost" type="button" onClick={onClose} disabled={mutation.isPending}>
            Отмена
          </Button>
          <Button type="submit" disabled={mutation.isPending}>
            {mutation.isPending ? "Перевод..." : "Перевести в производство"}
          </Button>
        </div>
      </form>
    </Modal>
  );
};
