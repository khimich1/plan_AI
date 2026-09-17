import { useEffect, useId, useMemo, useState, type KeyboardEvent } from "react";
import { useQuery } from "@tanstack/react-query";
import {
  searchCounterparties,
  type CounterpartyShort,
} from "@/features/commercial-offer/api/counterpartiesApi";
import { useDebouncedValue } from "@/shared/lib/useDebouncedValue";
import { Input } from "@/shared/ui/Field";

type Props = {
  selected: CounterpartyShort | null;
  onSelect: (item: CounterpartyShort | null) => void;
  onTextChange?: (value: string) => void;
  initialName?: string;
  placeholder?: string;
  disabled?: boolean;
};

const EMPTY_ITEMS: CounterpartyShort[] = [];

const normalizeName = (value: string): string => value.trim().toLowerCase();

export const formatCounterpartyRequisites = (item: Pick<CounterpartyShort, "inn" | "kpp">): string => {
  const inn = item.inn?.trim() ? item.inn.trim() : "—";
  const kpp = item.kpp?.trim() ? item.kpp.trim() : "—";
  return `ИНН ${inn} · КПП ${kpp}`;
};

export const resolveCounterpartyMatch = (
  text: string,
  items: CounterpartyShort[],
): CounterpartyShort | null => {
  const normalized = normalizeName(text);
  if (!normalized) {
    return null;
  }
  const exact = items.filter((item) => normalizeName(item.name) === normalized);
  return exact.length === 1 ? (exact[0] ?? null) : null;
};

export const CounterpartyAutocomplete = ({
  selected,
  onSelect,
  onTextChange,
  initialName = "",
  placeholder = "Начните вводить название или ИНН",
  disabled = false,
}: Props) => {
  const listId = useId();
  const optionIdPrefix = useId();
  const [text, setText] = useState(selected?.name ?? initialName);
  const [open, setOpen] = useState(false);
  const [highlightIndex, setHighlightIndex] = useState(0);
  const debounced = useDebouncedValue(text.trim());
  const canSearch = debounced.length >= 2;

  const query = useQuery({
    queryKey: ["counterparties", debounced],
    queryFn: () => searchCounterparties(debounced),
    enabled: canSearch,
  });

  const items = query.data?.items ?? EMPTY_ITEMS;

  useEffect(() => {
    if (selected) {
      setText(selected.name);
    }
  }, [selected?.id, selected?.name]);

  useEffect(() => {
    const match = resolveCounterpartyMatch(text, items);
    if (match && match.id !== selected?.id) {
      onSelect(match);
      setOpen(false);
    }
    // Search results arriving after the last keystroke can complete an exact match.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [items]);

  useEffect(() => {
    setHighlightIndex(0);
  }, [debounced, items.length]);

  const showList = open && canSearch;
  const activeOptionId =
    showList && items[highlightIndex] ? `${optionIdPrefix}-${items[highlightIndex].id}` : undefined;

  const handleChange = (value: string) => {
    setText(value);
    onTextChange?.(value);
    setOpen(true);
    if (!value.trim()) {
      if (selected) {
        onSelect(null);
      }
      return;
    }
    const match = resolveCounterpartyMatch(value, items);
    if (match) {
      if (match.id !== selected?.id) {
        onSelect(match);
      }
      return;
    }
    if (selected) {
      onSelect(null);
    }
  };

  const handleSelect = (item: CounterpartyShort) => {
    setText(item.name);
    onSelect(item);
    onTextChange?.(item.name);
    setOpen(false);
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === "Escape") {
      event.preventDefault();
      setOpen(false);
      return;
    }
    if (!open && (event.key === "ArrowDown" || event.key === "ArrowUp")) {
      setOpen(true);
      return;
    }
    if (event.key === "ArrowDown") {
      event.preventDefault();
      setHighlightIndex((index) => Math.min(index + 1, Math.max(items.length - 1, 0)));
      return;
    }
    if (event.key === "ArrowUp") {
      event.preventDefault();
      setHighlightIndex((index) => Math.max(index - 1, 0));
      return;
    }
    if (event.key === "Enter" && items[highlightIndex]) {
      event.preventDefault();
      handleSelect(items[highlightIndex]);
    }
  };

  const listLabel = useMemo(() => {
    if (query.isFetching) {
      return "Ищем…";
    }
    if (items.length === 0) {
      return "Не найдено среди клиентов";
    }
    return undefined;
  }, [items.length, query.isFetching]);

  return (
    <div style={{ position: "relative" }}>
      <Input
        type="text"
        role="combobox"
        aria-expanded={showList}
        aria-controls={listId}
        aria-activedescendant={activeOptionId}
        aria-autocomplete="list"
        aria-haspopup="listbox"
        value={text}
        onChange={(event) => handleChange(event.target.value)}
        onFocus={() => setOpen(true)}
        onBlur={() => setOpen(false)}
        onKeyDown={handleKeyDown}
        placeholder={placeholder}
        disabled={disabled}
        autoComplete="off"
      />
      {showList && (
        <ul
          id={listId}
          role="listbox"
          style={{
            position: "absolute",
            zIndex: 20,
            left: 0,
            right: 0,
            margin: "0.35rem 0 0",
            padding: "0.35rem 0",
            listStyle: "none",
            background: "#ffffff",
            border: "1px solid #d0d5dd",
            borderRadius: 12,
            boxShadow: "0 8px 24px rgba(15, 23, 42, 0.08)",
            maxHeight: 280,
            overflowY: "auto",
          }}
        >
          {listLabel && (
            <li
              role="presentation"
              style={{ padding: "0.65rem 0.9rem", color: "#667085", fontSize: "0.9rem" }}
            >
              {listLabel}
            </li>
          )}
          {items.map((item, index) => {
            const highlighted = index === highlightIndex;
            return (
              <li
                key={item.id}
                id={`${optionIdPrefix}-${item.id}`}
                role="option"
                aria-selected={highlighted}
                onMouseDown={(event) => event.preventDefault()}
                onMouseEnter={() => setHighlightIndex(index)}
                onClick={() => handleSelect(item)}
                style={{
                  padding: "0.65rem 0.9rem",
                  cursor: "pointer",
                  background: highlighted ? "#eef2ff" : "transparent",
                }}
              >
                <div style={{ fontWeight: 600 }}>{item.name}</div>
                <div style={{ color: "#475467", fontSize: "0.85rem", marginTop: "0.15rem" }}>
                  {formatCounterpartyRequisites(item)}
                </div>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
};
