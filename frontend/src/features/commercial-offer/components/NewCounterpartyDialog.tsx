import { useEffect, useState } from "react";
import {
  createCounterparty,
  DuplicateCounterpartyError,
  type CounterpartyShort,
} from "@/features/commercial-offer/api/counterpartiesApi";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Modal } from "@/shared/ui/Modal";
import { getErrorMessage } from "@/shared/lib/apiError";

type Props = {
  open: boolean;
  onClose: () => void;
  onCreated: (item: CounterpartyShort) => void;
  initialName?: string;
};

const emptyToNull = (value: string): string | null => {
  const trimmed = value.trim();
  return trimmed ? trimmed : null;
};

export const NewCounterpartyDialog = ({ open, onClose, onCreated, initialName = "" }: Props) => {
  const [name, setName] = useState(initialName);
  const [code1c, setCode1c] = useState("");
  const [inn, setInn] = useState("");
  const [kpp, setKpp] = useState("");
  const [isClient, setIsClient] = useState(true);
  const [pending, setPending] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [warning, setWarning] = useState<string | null>(null);
  const [existing, setExisting] = useState<CounterpartyShort | null>(null);

  useEffect(() => {
    if (!open) {
      return;
    }
    setName(initialName);
    setCode1c("");
    setInn("");
    setKpp("");
    setIsClient(true);
    setPending(false);
    setError(null);
    setWarning(null);
    setExisting(null);
    // Seed once when the dialog opens; ignore later initialName updates (e.g. parent select).
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const canSubmit = Boolean(name.trim() && code1c.trim()) && !pending;

  const handleSubmit = async () => {
    if (!canSubmit) {
      return;
    }
    setPending(true);
    setError(null);
    setWarning(null);
    setExisting(null);
    try {
      const result = await createCounterparty({
        name: name.trim(),
        code_1c: code1c.trim(),
        inn: emptyToNull(inn),
        kpp: emptyToNull(kpp),
        is_client: isClient,
      });
      if (result.warning) {
        setWarning(result.warning);
      }
      onCreated(result.item);
      if (!result.warning) {
        onClose();
      }
    } catch (err) {
      if (err instanceof DuplicateCounterpartyError) {
        setExisting(err.existing);
        setError(err.message);
      } else {
        setError(getErrorMessage(err));
      }
    } finally {
      setPending(false);
    }
  };

  return (
    <Modal open={open} onClose={onClose} title="Новый контрагент" maxWidth={480}>
      <div style={{ display: "grid", gap: "0.9rem" }}>
        <Alert tone="info">Сначала заведите контрагента в 1С и скопируйте код.</Alert>
        {error && <Alert tone="error">{error}</Alert>}
        {warning && <Alert tone="info">{warning}</Alert>}
        {existing && (
          <div
            style={{
              border: "1px solid #e4e7ec",
              borderRadius: 12,
              padding: "0.75rem 0.9rem",
              display: "grid",
              gap: "0.5rem",
            }}
          >
            <strong>{existing.name}</strong>
            <span style={{ color: "#475467", fontSize: "0.9rem" }}>
              код 1С {existing.code_1c}
              {existing.inn ? ` · ИНН ${existing.inn}` : ""}
            </span>
            <Button
              type="button"
              variant="secondary"
              onClick={() => {
                onCreated(existing);
                onClose();
              }}
            >
              Выбрать её
            </Button>
          </div>
        )}
        <FieldWrapper label="Наименование">
          <Input value={name} onChange={(event) => setName(event.target.value)} autoComplete="off" />
        </FieldWrapper>
        <FieldWrapper label="Код 1С">
          <Input value={code1c} onChange={(event) => setCode1c(event.target.value)} autoComplete="off" />
        </FieldWrapper>
        <FieldWrapper label="ИНН">
          <Input value={inn} onChange={(event) => setInn(event.target.value)} autoComplete="off" />
        </FieldWrapper>
        <FieldWrapper label="КПП">
          <Input value={kpp} onChange={(event) => setKpp(event.target.value)} autoComplete="off" />
        </FieldWrapper>
        <label style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
          <input
            type="checkbox"
            checked={isClient}
            onChange={(event) => setIsClient(event.target.checked)}
          />
          Клиент
        </label>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
          <Button type="button" variant="ghost" onClick={onClose} disabled={pending}>
            Отмена
          </Button>
          <Button type="button" onClick={() => void handleSubmit()} disabled={!canSubmit}>
            {pending ? "Добавляем..." : "Добавить"}
          </Button>
        </div>
      </div>
    </Modal>
  );
};
