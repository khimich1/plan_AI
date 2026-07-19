import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { Alert } from "@/shared/ui/Alert";
import { useDeleteOfferMutation } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { getErrorMessage } from "@/shared/lib/apiError";

type Props = {
  open: boolean;
  onClose: () => void;
  onDeleted: () => void;
  kpId: number;
  customerName: string | null;
};

export const DeleteConfirmDialog = ({ open, onClose, onDeleted, kpId, customerName }: Props) => {
  const mutation = useDeleteOfferMutation();

  const onConfirm = async () => {
    try {
      await mutation.mutateAsync(kpId);
      onDeleted();
      onClose();
    } catch {
      // ошибка показывается через mutation.error
    }
  };

  return (
    <Modal open={open} onClose={onClose} title={`Удалить КП №${kpId}?`}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <Alert tone="warning">
          Действие необратимо. Будут удалены информация о КП, список плит, файлы и метаданные.
        </Alert>
        {customerName && (
          <div>
            <strong>Клиент:</strong> {customerName}
          </div>
        )}
        {mutation.isError && <Alert tone="error">{getErrorMessage(mutation.error)}</Alert>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button variant="ghost" onClick={onClose} disabled={mutation.isPending}>
            Отмена
          </Button>
          <Button variant="danger" onClick={onConfirm} disabled={mutation.isPending}>
            {mutation.isPending ? "Удаление..." : "Да, удалить"}
          </Button>
        </div>
      </div>
    </Modal>
  );
};
