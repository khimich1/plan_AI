import { useEffect, useState } from "react";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";
import { Alert } from "@/shared/ui/Alert";
import type { ArchiveSearchState } from "@/features/commercial-archive/types/archive";

type Props = {
  onSubmit: (state: ArchiveSearchState) => void;
  onClear: () => void;
  activeQuery: ArchiveSearchState;
};

export const ArchiveSearchBar = ({ onSubmit, onClear, activeQuery }: Props) => {
  const [kpIdValue, setKpIdValue] = useState<string>(
    activeQuery?.kind === "number" ? String(activeQuery.value) : "",
  );
  const [customerValue, setCustomerValue] = useState<string>(
    activeQuery?.kind === "customer" ? activeQuery.value : "",
  );
  const [validationError, setValidationError] = useState<string | null>(null);

  useEffect(() => {
    if (activeQuery === null) {
      setKpIdValue("");
      setCustomerValue("");
      setValidationError(null);
      return;
    }
    if (activeQuery.kind === "number") {
      setKpIdValue(String(activeQuery.value));
      setCustomerValue("");
    } else {
      setCustomerValue(activeQuery.value);
      setKpIdValue("");
    }
    setValidationError(null);
  }, [activeQuery]);

  const handleSubmit = (event: React.FormEvent) => {
    event.preventDefault();
    setValidationError(null);

    const kpIdRaw = kpIdValue.trim();
    const customerRaw = customerValue.trim();

    if (kpIdRaw) {
      const kpId = Number(kpIdRaw);
      if (!Number.isFinite(kpId) || kpId <= 0 || !Number.isInteger(kpId)) {
        setValidationError("Укажите корректный номер КП (целое число больше 0).");
        return;
      }
      onSubmit({ kind: "number", value: kpId });
      return;
    }

    if (customerRaw) {
      if (customerRaw.length < 2) {
        setValidationError("Имя заказчика должно содержать не менее 2 символов.");
        return;
      }
      onSubmit({ kind: "customer", value: customerRaw });
      return;
    }

    setValidationError("Укажите номер КП или имя заказчика.");
  };

  const handleClear = () => {
    setKpIdValue("");
    setCustomerValue("");
    setValidationError(null);
    onClear();
  };

  return (
    <div style={{ display: "grid", gap: "0.5rem" }}>
      <form
        onSubmit={handleSubmit}
        style={{ display: "flex", gap: "0.5rem", alignItems: "center", flexWrap: "wrap" }}
      >
        <div style={{ width: 140, flex: "0 0 140px" }}>
          <Input
            type="number"
            min={1}
            placeholder="Номер КП"
            value={kpIdValue}
            onChange={(event) => setKpIdValue(event.target.value)}
          />
        </div>
        <div style={{ minWidth: 200, flex: "1 1 200px" }}>
          <Input
            type="text"
            placeholder="Название заказчика"
            value={customerValue}
            onChange={(event) => setCustomerValue(event.target.value)}
          />
        </div>
        <Button type="submit" variant="secondary">
          Найти
        </Button>
        {activeQuery !== null && (
          <Button type="button" variant="ghost" onClick={handleClear}>
            Сбросить
          </Button>
        )}
      </form>
      {validationError && <Alert tone="warning">{validationError}</Alert>}
    </div>
  );
};
