import { useState } from "react";
import { Button } from "@/shared/ui/Button";
import { Input } from "@/shared/ui/Field";

type Props = {
  onSubmit: (kpId: number) => void;
  onClear: () => void;
  activeQuery: number | null;
};

export const ArchiveSearchBar = ({ onSubmit, onClear, activeQuery }: Props) => {
  const [value, setValue] = useState<string>(activeQuery ? String(activeQuery) : "");

  const handleSubmit = (event: React.FormEvent) => {
    event.preventDefault();
    const kpId = Number(value.trim());
    if (!Number.isFinite(kpId) || kpId <= 0) {
      return;
    }
    onSubmit(kpId);
  };

  const handleClear = () => {
    setValue("");
    onClear();
  };

  return (
    <form onSubmit={handleSubmit} style={{ display: "flex", gap: "0.5rem", alignItems: "center", flexWrap: "wrap" }}>
      <div style={{ minWidth: 220, flex: "1 0 220px" }}>
        <Input
          type="number"
          min={1}
          placeholder="Найти по номеру КП"
          value={value}
          onChange={(event) => setValue(event.target.value)}
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
  );
};
