import { useEffect, useId, useState } from "react";
import { Input } from "@/shared/ui/Field";
import { useCarriersQuery } from "@/features/logistics/hooks/useLogisticsQueries";
import { useDebouncedValue } from "@/features/logistics/lib/useDebouncedValue";
import type { Carrier } from "@/features/logistics/types/logistics";

type Props = {
  selected: { id: number; name: string } | null;
  onSelect: (carrier: { id: number; name: string } | null) => void;
  placeholder?: string;
  disabled?: boolean;
};

export const resolveCarrierMatch = (
  text: string,
  carriers: Carrier[],
): { id: number; name: string } | null => {
  const normalized = text.trim().toLowerCase();
  if (!normalized) {
    return null;
  }
  const exact = carriers.find((c) => c.name.trim().toLowerCase() === normalized);
  return exact ? { id: exact.id, name: exact.name } : null;
};

export const CarrierAutocomplete = ({
  selected,
  onSelect,
  placeholder = "Начните вводить название",
  disabled = false,
}: Props) => {
  const listId = useId();
  const [text, setText] = useState(selected?.name ?? "");
  const debounced = useDebouncedValue(text.trim());
  const carriersQuery = useCarriersQuery({
    q: debounced.length >= 2 ? debounced : "",
    activeOnly: true,
  });
  const carriers = carriersQuery.data ?? [];

  useEffect(() => {
    setText(selected?.name ?? "");
  }, [selected?.id, selected?.name]);

  // Точное совпадение может прийти с debounced-запросом после последнего keystroke.
  useEffect(() => {
    const match = resolveCarrierMatch(text, carriers);
    if (match && match.id !== selected?.id) {
      onSelect(match);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [carriers]);

  const handleChange = (value: string) => {
    setText(value);
    if (!value.trim()) {
      onSelect(null);
      return;
    }
    onSelect(resolveCarrierMatch(value, carriers));
  };

  return (
    <>
      <Input
        type="text"
        value={text}
        onChange={(event) => handleChange(event.target.value)}
        placeholder={placeholder}
        list={listId}
        disabled={disabled}
        autoComplete="off"
      />
      <datalist id={listId}>
        {carriers.map((carrier) => (
          <option key={carrier.id} value={carrier.name} />
        ))}
      </datalist>
    </>
  );
};
