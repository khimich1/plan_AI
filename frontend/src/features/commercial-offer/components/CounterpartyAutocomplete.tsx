import { useEffect, useId, useMemo, useState, type KeyboardEvent } from "react";
import { useQuery } from "@tanstack/react-query";
import {
  searchContractCards,
  searchCounterparties,
  type ContractCardHit,
  type CounterpartyShort,
} from "@/features/commercial-offer/api/counterpartiesApi";
import { useDebouncedValue } from "@/shared/lib/useDebouncedValue";
import { Input } from "@/shared/ui/Field";

export type CounterpartySearchSource = "directory" | "contract-cards";

type Props = {
  selected: CounterpartyShort | null;
  onSelect: (item: CounterpartyShort | null) => void;
  onTextChange?: (value: string) => void;
  onContractCardSelect?: (card: ContractCardHit) => void;
  initialName?: string;
  placeholder?: string;
  disabled?: boolean;
  searchSource?: CounterpartySearchSource;
};

const EMPTY_ITEMS: CounterpartyShort[] = [];
const EMPTY_CARDS: ContractCardHit[] = [];

const normalizeName = (value: string): string => value.trim().toLowerCase();

export const formatCounterpartyRequisites = (item: Pick<CounterpartyShort, "inn" | "kpp">): string => {
  const inn = item.inn?.trim() ? item.inn.trim() : "—";
  const kpp = item.kpp?.trim() ? item.kpp.trim() : "—";
  return `ИНН ${inn} · КПП ${kpp}`;
};

const formatDirectoryDetail = (item: CounterpartyShort): string => {
  const code = item.code_1c?.trim() ? item.code_1c.trim() : "—";
  return `${formatCounterpartyRequisites(item)} · код ${code}`;
};

const formatContractCardDetail = (card: ContractCardHit): string => {
  const inn = card.inn?.trim() ? card.inn.trim() : "—";
  const incomplete = card.requisites_complete ? "" : " · не полные";
  return `ИНН ${inn} · ${card.number}${incomplete}`;
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

const uniqueByName = <T extends { name: string }>(text: string, rows: T[]): T | null => {
  const normalized = normalizeName(text);
  if (!normalized) {
    return null;
  }
  const exact = rows.filter((row) => normalizeName(row.name) === normalized);
  return exact.length === 1 ? (exact[0] ?? null) : null;
};

const pickFromCard = (card: ContractCardHit): CounterpartyShort | null => {
  if (!card.counterparty) {
    return null;
  }
  return {
    id: card.counterparty.id,
    code_1c: card.counterparty.code_1c,
    name: card.name,
    inn: card.inn,
    kpp: card.counterparty.kpp,
  };
};

type SuggestionRow = {
  key: string;
  label: string;
  detail: string;
  directoryItem: CounterpartyShort | null;
  card: ContractCardHit | null;
};

export const CounterpartyAutocomplete = ({
  selected,
  onSelect,
  onTextChange,
  onContractCardSelect,
  initialName = "",
  placeholder = "Начните вводить название или ИНН",
  disabled = false,
  searchSource = "directory",
}: Props) => {
  const listId = useId();
  const optionIdPrefix = useId();
  const [text, setText] = useState(selected?.name ?? initialName);
  const [open, setOpen] = useState(false);
  const [highlightIndex, setHighlightIndex] = useState(0);
  const debounced = useDebouncedValue(text.trim());
  const canSearch = debounced.length >= 2;
  const cardsMode = searchSource === "contract-cards";

  const directoryQuery = useQuery({
    queryKey: ["counterparties", debounced],
    queryFn: () => searchCounterparties(debounced),
    enabled: canSearch && !cardsMode,
  });
  const cardQuery = useQuery({
    queryKey: ["counterparties", "contract-cards", debounced],
    queryFn: () => searchContractCards(debounced),
    enabled: canSearch && cardsMode,
  });

  const items = cardsMode ? EMPTY_ITEMS : (directoryQuery.data?.items ?? EMPTY_ITEMS);
  const cards = cardsMode ? (cardQuery.data?.items ?? EMPTY_CARDS) : EMPTY_CARDS;
  const isFetching = cardsMode ? cardQuery.isFetching : directoryQuery.isFetching;

  const applyCard = (card: ContractCardHit) => {
    const linked = pickFromCard(card);
    setText(card.name);
    onTextChange?.(card.name);
    if (linked) {
      onSelect(linked);
    } else {
      onContractCardSelect?.(card);
    }
    setOpen(false);
  };

  useEffect(() => {
    if (selected) {
      setText(selected.name);
    }
  }, [selected?.id, selected?.name]);

  useEffect(() => {
    if (cardsMode) {
      return;
    }
    const match = resolveCounterpartyMatch(text, items);
    if (match && match.id !== selected?.id) {
      onSelect(match);
      setOpen(false);
    }
    // Search results arriving after the last keystroke can complete an exact match.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [items]);

  useEffect(() => {
    if (!cardsMode) {
      return;
    }
    const match = uniqueByName(text, cards);
    if (!match) {
      return;
    }
    applyCard(match);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [cards]);

  const rows = useMemo<SuggestionRow[]>(() => {
    if (cardsMode) {
      return cards.map((card) => ({
        key: `card-${card.id}`,
        label: card.name,
        detail: formatContractCardDetail(card),
        directoryItem: null,
        card,
      }));
    }
    return items.map((item) => ({
      key: `cp-${item.id}`,
      label: item.name,
      detail: formatDirectoryDetail(item),
      directoryItem: item,
      card: null,
    }));
  }, [cards, cardsMode, items]);

  useEffect(() => {
    setHighlightIndex(0);
  }, [debounced, rows.length]);

  const showList = open && canSearch;
  const activeOptionId =
    showList && rows[highlightIndex] ? `${optionIdPrefix}-${rows[highlightIndex].key}` : undefined;

  const chooseRow = (row: SuggestionRow) => {
    if (row.card) {
      applyCard(row.card);
      return;
    }
    if (!row.directoryItem) {
      return;
    }
    setText(row.directoryItem.name);
    onSelect(row.directoryItem);
    onTextChange?.(row.directoryItem.name);
    setOpen(false);
  };

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
    if (cardsMode) {
      const match = uniqueByName(value, cards);
      if (match) {
        applyCard(match);
        return;
      }
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
      setHighlightIndex((index) => Math.min(index + 1, Math.max(rows.length - 1, 0)));
      return;
    }
    if (event.key === "ArrowUp") {
      event.preventDefault();
      setHighlightIndex((index) => Math.max(index - 1, 0));
      return;
    }
    if (event.key === "Enter" && rows[highlightIndex]) {
      event.preventDefault();
      chooseRow(rows[highlightIndex]);
    }
  };

  const listLabel = useMemo(() => {
    if (isFetching) {
      return "Ищем…";
    }
    if (rows.length === 0) {
      return cardsMode ? "Не найдено среди карточек" : "Не найдено среди клиентов";
    }
    return undefined;
  }, [cardsMode, isFetching, rows.length]);

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
          {rows.map((row, index) => {
            const highlighted = index === highlightIndex;
            return (
              <li
                key={row.key}
                id={`${optionIdPrefix}-${row.key}`}
                role="option"
                aria-selected={highlighted}
                onMouseDown={(event) => event.preventDefault()}
                onMouseEnter={() => setHighlightIndex(index)}
                onClick={() => chooseRow(row)}
                style={{
                  padding: "0.65rem 0.9rem",
                  cursor: "pointer",
                  background: highlighted ? "#eef2ff" : "transparent",
                }}
              >
                <div style={{ fontWeight: 600 }}>{row.label}</div>
                <div style={{ color: "#475467", fontSize: "0.85rem", marginTop: "0.15rem" }}>
                  {row.detail}
                </div>
              </li>
            );
          })}
        </ul>
      )}
    </div>
  );
};
