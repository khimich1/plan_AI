import { useEffect, useState, type ReactNode } from "react";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { Alert } from "@/shared/ui/Alert";
import { getErrorMessage } from "@/shared/lib/apiError";

type Props = {
  open: boolean;
  onClose: () => void;
  title: string;
  description: ReactNode;
  confirmLabel: string;
  isPending: boolean;
  isError: boolean;
  error: unknown;
  onConfirm: () => void;
  /** Если задано (например, "ОБНУЛИТЬ"), пользователь должен ввести это слово,
   *  чтобы разблокировать кнопку «Подтвердить». Используется для
   *  особо опасных операций (полное обнуление). */
  confirmKeyword?: string;
};

export const ResetConfirmDialog = ({
  open,
  onClose,
  title,
  description,
  confirmLabel,
  isPending,
  isError,
  error,
  onConfirm,
  confirmKeyword,
}: Props) => {
  const [typedKeyword, setTypedKeyword] = useState("");

  useEffect(() => {
    if (!open) {
      setTypedKeyword("");
    }
  }, [open]);

  const keywordOk =
    !confirmKeyword || typedKeyword.trim() === confirmKeyword;
  const disabled = isPending || !keywordOk;

  return (
    <Modal open={open} onClose={onClose} title={title} maxWidth={520}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <Alert tone="warning">{description}</Alert>

        {confirmKeyword && (
          <label style={{ display: "grid", gap: "0.5rem" }}>
            <span style={{ fontSize: "0.9rem", color: "#475467" }}>
              Введите <code>{confirmKeyword}</code>, чтобы подтвердить:
            </span>
            <input
              type="text"
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
        )}

        {isError && <Alert tone="error">{getErrorMessage(error)}</Alert>}

        <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            gap: "0.5rem",
          }}
        >
          <Button variant="ghost" onClick={onClose} disabled={isPending}>
            Отмена
          </Button>
          <Button variant="danger" onClick={onConfirm} disabled={disabled}>
            {isPending ? "Выполняется..." : confirmLabel}
          </Button>
        </div>
      </div>
    </Modal>
  );
};
