import { useEffect, useState } from "react";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Alert } from "@/shared/ui/Alert";
import { useUpdateDiscountMutation } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { getErrorMessage } from "@/shared/lib/apiError";

type Props = {
  open: boolean;
  onClose: () => void;
  kpId: number;
  currentDiscount: number;
};

export const DiscountEditDialog = ({ open, onClose, kpId, currentDiscount }: Props) => {
  const [value, setValue] = useState<string>("");
  const [inputError, setInputError] = useState<string>();
  const mutation = useUpdateDiscountMutation();

  useEffect(() => {
    if (open) {
      setValue(String(currentDiscount ?? 0));
      setInputError(undefined);
    }
  }, [open, currentDiscount]);

  useEffect(() => {
    if (open) {
      mutation.reset();
    }
  }, [open]);

  const onSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    const normalized = value.trim().replace(",", ".");
    if (normalized.length === 0) {
      setInputError("Введите значение скидки");
      return;
    }
    const parsed = Number(normalized);
    if (!Number.isFinite(parsed) || parsed < 0 || parsed > 100) {
      setInputError("Введите значение от 0 до 100");
      return;
    }
    setInputError(undefined);
    try {
      await mutation.mutateAsync({ kpId, discount: parsed });
      onClose();
    } catch {
      // ошибка показывается через mutation.error
    }
  };

  return (
    <Modal open={open} onClose={onClose} title={`Скидка для КП №${kpId}`}>
      <form onSubmit={onSubmit} style={{ display: "grid", gap: "1rem" }}>
        <FieldWrapper label="Новый процент скидки" hint="Значение от 0 до 100" error={inputError}>
          <Input
            type="number"
            min={0}
            max={100}
            step="0.1"
            value={value}
            onChange={(event) => {
              setValue(event.target.value);
              if (inputError) {
                setInputError(undefined);
              }
            }}
            autoFocus
          />
        </FieldWrapper>
        {mutation.isError && <Alert tone="error">{getErrorMessage(mutation.error)}</Alert>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button variant="ghost" type="button" onClick={onClose} disabled={mutation.isPending}>
            Отмена
          </Button>
          <Button type="submit" disabled={mutation.isPending}>
            {mutation.isPending ? "Сохранение..." : "Сохранить"}
          </Button>
        </div>
      </form>
    </Modal>
  );
};
