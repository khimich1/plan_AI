import { useEffect, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Modal } from "@/shared/ui/Modal";
import {
  HIGH_DISCOUNT_CONFIRMATION_KEYWORD,
  HIGH_DISCOUNT_WARNING,
} from "@/features/commercial-offer/lib/discountFromTargetSum";

type HighDiscountConfirmDialogProps = {
  open: boolean;
  discountPercent: number;
  isPending: boolean;
  onConfirm: () => void;
  onCancel: () => void;
};

export const HighDiscountConfirmDialog = ({
  open,
  discountPercent,
  isPending,
  onConfirm,
  onCancel,
}: HighDiscountConfirmDialogProps) => {
  const [typedKeyword, setTypedKeyword] = useState("");

  useEffect(() => {
    if (!open) {
      setTypedKeyword("");
    }
  }, [open]);

  const confirmed = typedKeyword.trim() === HIGH_DISCOUNT_CONFIRMATION_KEYWORD;

  return (
    <Modal open={open} onClose={onCancel} title="Подтверждение скидки" maxWidth={520}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <Alert tone="warning">
          {HIGH_DISCOUNT_WARNING} Текущая скидка: {discountPercent.toFixed(2)}%.
        </Alert>
        <label style={{ display: "grid", gap: "0.5rem" }}>
          <span style={{ fontSize: "0.9rem", color: "#475467" }}>
            Введите <code>{HIGH_DISCOUNT_CONFIRMATION_KEYWORD}</code>, чтобы подтвердить:
          </span>
          <input
            value={typedKeyword}
            onChange={(event) => setTypedKeyword(event.target.value)}
            autoFocus
            disabled={isPending}
            style={{
              padding: "0.6rem 0.8rem",
              borderRadius: 10,
              border: "1px solid #d6defa",
              fontSize: "0.95rem",
              fontFamily: "inherit",
            }}
          />
        </label>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button variant="ghost" onClick={onCancel} disabled={isPending}>
            Отмена
          </Button>
          <Button onClick={onConfirm} disabled={isPending || !confirmed}>
            {isPending ? "Применяем..." : "OK"}
          </Button>
        </div>
      </div>
    </Modal>
  );
};
