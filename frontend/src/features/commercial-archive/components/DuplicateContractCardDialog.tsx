import { Button } from "@/shared/ui/Button";
import { Modal } from "@/shared/ui/Modal";
import type { ContractCard } from "@/features/commercial-archive/types/supplyContract";

type DuplicateContractCardDialogProps = {
  items: ContractCard[];
  busy: boolean;
  onEdit: (card: ContractCard) => void;
  onCreateAnyway: () => void;
  onClose: () => void;
};

const shownInn = (inn: string | null) => inn?.trim() || "—";

export const DuplicateContractCardDialog = ({
  items,
  busy,
  onEdit,
  onCreateAnyway,
  onClose,
}: DuplicateContractCardDialogProps) => {
  if (items.length === 0) {
    return null;
  }

  return (
    <Modal open onClose={onClose} title="Такая карточка уже есть">
      <ul
        aria-label="Найденные карточки"
        style={{ listStyle: "none", margin: 0, padding: 0, display: "grid", gap: "0.5rem" }}
      >
        {items.map((item, index) => (
          <li
            key={item.id}
            style={{
              display: "grid",
              gap: "0.35rem",
              padding: "0.75rem 0",
              borderTop: "1px solid #e4e7ec",
            }}
          >
            <span style={{ fontWeight: 600, overflowWrap: "anywhere" }}>{item.name}</span>
            <span style={{ color: "#475467" }}>{shownInn(item.inn)}</span>
            <span>{item.number}</span>
            <div>
              <Button
                type="button"
                variant="secondary"
                autoFocus={index === 0}
                disabled={busy}
                onClick={() => onEdit(item)}
              >
                Редактировать
              </Button>
            </div>
          </li>
        ))}
      </ul>
      <div style={{ marginTop: "1rem" }}>
        <Button type="button" disabled={busy} onClick={onCreateAnyway}>
          Всё равно создать
        </Button>
      </div>
    </Modal>
  );
};
