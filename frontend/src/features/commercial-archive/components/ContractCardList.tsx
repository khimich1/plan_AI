import { useEffect, useRef, useState, type CSSProperties, type DragEvent, type FormEvent } from "react";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import {
  useContractCardsQuery,
  useCreateContractCardMutation,
  useParseContractCardFileMutation,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import type {
  ContractCard,
  ContractCardCreatePayload,
  SupplyContractParseResult,
} from "@/features/commercial-archive/types/supplyContract";
import { ContractCardDrawer, type ContractCardIntake } from "./ContractCardDrawer";
import { DuplicateContractCardDialog } from "./DuplicateContractCardDialog";
import { getErrorMessage } from "@/shared/lib/apiError";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input } from "@/shared/ui/Field";

const MIN_SEARCH_LENGTH = 2;
const COLUMN_GRID = "minmax(0, 1.6fr) minmax(8.5rem, 0.9fr) minmax(8rem, 0.8fr)";
const CARD_FILE_ACCEPT = ".docx,.pdf,.xlsx,image/jpeg,image/png,application/pdf";

const buildCreatePayload = (
  name: string,
  inn: string,
  number: string,
): ContractCardCreatePayload | null => {
  const trimmedName = name.trim();
  if (!trimmedName) {
    return null;
  }
  const payload: ContractCardCreatePayload = { name: trimmedName };
  const trimmedInn = inn.trim();
  const trimmedNumber = number.trim();
  if (trimmedInn) {
    payload.inn = trimmedInn;
  }
  if (trimmedNumber) {
    payload.number = trimmedNumber;
  }
  return payload;
};

const trimmedText = (value: string | null | undefined) => value?.trim() ?? "";

const fileStem = (filename: string) => {
  const base = filename.split(/[/\\]/).pop() ?? filename;
  const dot = base.lastIndexOf(".");
  if (dot <= 0) {
    return base.trim();
  }
  return base.slice(0, dot).trim();
};

const nameFromParsedCard = (parsed: SupplyContractParseResult, filename: string) => {
  const shortName = trimmedText(parsed.fields.short_name);
  if (shortName) {
    return shortName;
  }
  const fullName = trimmedText(parsed.fields.full_name);
  if (fullName) {
    return fullName;
  }
  return fileStem(filename);
};

const payloadFromParsedCard = (
  parsed: SupplyContractParseResult,
  filename: string,
): ContractCardCreatePayload | null => {
  const name = nameFromParsedCard(parsed, filename);
  if (!name) {
    return null;
  }
  const inn = trimmedText(parsed.fields.inn);
  return inn ? { name, inn } : { name };
};

const listInn = (inn: string | null): string => inn?.trim() || "—";

const openedInn = (inn: string | null): string => {
  const trimmed = inn?.trim();
  return trimmed ? `ИНН ${trimmed}` : "ИНН не указан";
};

const rowStyle = (selected: boolean): CSSProperties => ({
  display: "grid",
  gridTemplateColumns: COLUMN_GRID,
  gap: "0.75rem",
  alignItems: "center",
  width: "100%",
  padding: "0.9rem 1rem",
  border: `1px solid ${selected ? "#2b5cff" : "#e4e7ec"}`,
  borderRadius: 14,
  background: selected ? "#eef2ff" : "#ffffff",
  cursor: "pointer",
  textAlign: "left",
  font: "inherit",
  color: "inherit",
});

const OpenedContractCard = ({
  card,
  intake,
  onUpdated,
  onDeleted,
}: {
  card: ContractCard;
  intake: ContractCardIntake | null;
  onUpdated: (next: ContractCard) => void;
  onDeleted: () => void;
}) => {
  const ref = useRef<HTMLElement>(null);

  useEffect(() => {
    ref.current?.focus();
  }, [card.id]);

  return (
    <section ref={ref} tabIndex={-1} aria-label="Карточка договора">
      <Card title={card.name}>
        <p style={{ margin: 0 }}>
          Номер <span>{card.number}</span>
        </p>
        <p style={{ margin: "0.35rem 0 0", color: "#475467" }}>{openedInn(card.inn)}</p>
        <ContractCardDrawer card={card} intake={intake} onUpdated={onUpdated} onDeleted={onDeleted} />
      </Card>
    </section>
  );
};

const ContractCardRows = ({
  items,
  selectedId,
  onSelect,
}: {
  items: ContractCard[];
  selectedId: number | null;
  onSelect: (card: ContractCard) => void;
}) => (
  <div style={{ display: "grid", gap: "0.5rem" }}>
    <div
      aria-hidden="true"
      style={{
        display: "grid",
        gridTemplateColumns: COLUMN_GRID,
        gap: "0.75rem",
        padding: "0 1rem",
        color: "#475467",
        fontSize: "0.85rem",
        fontWeight: 600,
      }}
    >
      <span>Имя</span>
      <span>ИНН</span>
      <span>Номер</span>
    </div>
    <ul aria-label="Карточки договоров" style={{ listStyle: "none", margin: 0, padding: 0, display: "grid", gap: "0.5rem" }}>
      {items.map((item) => {
        const selected = item.id === selectedId;
        return (
          <li key={item.id}>
            <button
              type="button"
              aria-pressed={selected}
              onClick={() => onSelect(item)}
              style={rowStyle(selected)}
            >
              <span style={{ fontWeight: 600, overflowWrap: "anywhere" }}>
                {item.name}
                {item.requisites_complete === false ? (
                  <span style={{ marginLeft: "0.5rem", color: "#b54708", fontWeight: 600 }}>не полные</span>
                ) : null}
              </span>
              <span>{listInn(item.inn)}</span>
              <span>{item.number}</span>
            </button>
          </li>
        );
      })}
    </ul>
  </div>
);

const ContractCardResults = ({
  query,
  selectedId,
  onSelect,
}: {
  query: string;
  selectedId: number | null;
  onSelect: (card: ContractCard) => void;
}) => {
  const search = useContractCardsQuery(query);

  if (search.isError) {
    return <Alert tone="error">{getErrorMessage(search.error)}</Alert>;
  }
  if (!search.data) {
    return (
      <p role="status" style={{ margin: 0, color: "#475467" }}>
        Ищу карточки…
      </p>
    );
  }
  if (search.data.items.length === 0) {
    return (
      <p role="status" style={{ margin: 0, color: "#475467" }}>
        {query ? "Ничего не найдено" : "Карточек пока нет"}
      </p>
    );
  }
  return <ContractCardRows items={search.data.items} selectedId={selectedId} onSelect={onSelect} />;
};

const NewCardFileZone = ({
  disabled,
  onFile,
}: {
  disabled: boolean;
  onFile: (file: File | undefined) => void;
}) => {
  const [dragOver, setDragOver] = useState(false);

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragOver(false);
    if (disabled) {
      return;
    }
    onFile(event.dataTransfer.files?.[0]);
  };

  return (
    <div
      aria-label="Зона файла новой карточки"
      onDragEnter={(event) => {
        event.preventDefault();
        setDragOver(true);
      }}
      onDragOver={(event) => {
        event.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={(event) => {
        event.preventDefault();
        setDragOver(false);
      }}
      onDrop={onDrop}
      style={{
        border: `2px dashed ${dragOver ? "#2b5cff" : "#d0d5dd"}`,
        borderRadius: 14,
        padding: "1rem 1.25rem",
        background: dragOver ? "#eef2ff" : "#f8faff",
        color: "#344054",
      }}
    >
      <FieldWrapper label="Карточка контрагента">
        <input
          aria-label="Файл новой карточки"
          type="file"
          accept={CARD_FILE_ACCEPT}
          disabled={disabled}
          onChange={(event) => onFile(event.target.files?.[0])}
        />
      </FieldWrapper>
      <div style={{ fontSize: "0.9rem", color: "#667085" }}>или перетащите файл сюда</div>
    </div>
  );
};

type PendingCardCreate = {
  payload: ContractCardCreatePayload;
  intake: ContractCardIntake | null;
};

const NewContractCardForm = ({
  onCreated,
}: {
  onCreated: (card: ContractCard, intake?: ContractCardIntake | null) => void;
}) => {
  const create = useCreateContractCardMutation();
  const parseFile = useParseContractCardFileMutation();
  const busyRef = useRef(false);
  const [name, setName] = useState("");
  const [inn, setInn] = useState("");
  const [number, setNumber] = useState("");
  const [fileError, setFileError] = useState<string | null>(null);
  const [duplicates, setDuplicates] = useState<ContractCard[]>([]);
  const [pending, setPending] = useState<PendingCardCreate | null>(null);
  const [checking, setChecking] = useState(false);
  const busy = create.isPending || parseFile.isPending || checking;
  const paused = duplicates.length > 0;

  const clearForm = () => {
    setName("");
    setInn("");
    setNumber("");
  };

  const dismissDuplicates = () => {
    setDuplicates([]);
    setPending(null);
  };

  const editDuplicate = (found: ContractCard) => {
    setDuplicates([]);
    setPending(null);
    onCreated(found);
  };

  const createAnyway = async () => {
    if (!pending || busyRef.current) {
      return;
    }
    busyRef.current = true;
    const current = pending;
    setFileError(null);
    try {
      const created = await create.mutateAsync(current.payload);
      clearForm();
      setDuplicates([]);
      setPending(null);
      onCreated(created, current.intake);
    } catch (error) {
      setDuplicates([]);
      setPending(null);
      setFileError(getErrorMessage(error));
    } finally {
      busyRef.current = false;
    }
  };

  const rememberOrCreate = async (
    payload: ContractCardCreatePayload,
    intake: ContractCardIntake | null,
  ) => {
    const found = await archiveApi.findDuplicateContractCards(payload.name, payload.inn ?? "");
    if (found.items.length > 0) {
      setPending({ payload, intake });
      setDuplicates(found.items);
      return;
    }
    const created = await create.mutateAsync(payload);
    clearForm();
    onCreated(created, intake);
  };

  const guard = async (action: () => Promise<void>) => {
    if (busyRef.current || paused) {
      return;
    }
    busyRef.current = true;
    setChecking(true);
    setFileError(null);
    try {
      await action();
    } catch (error) {
      setFileError(getErrorMessage(error));
    } finally {
      busyRef.current = false;
      setChecking(false);
    }
  };

  const handleSubmit = (event: FormEvent) => {
    event.preventDefault();
    const payload = buildCreatePayload(name, inn, number);
    if (!payload) {
      return;
    }
    void guard(() => rememberOrCreate(payload, null));
  };

  const acceptFile = (file: File | undefined) => {
    if (!file) {
      return;
    }
    void guard(async () => {
      const parsed = await parseFile.mutateAsync(file);
      const payload = payloadFromParsedCard(parsed, file.name);
      if (!payload) {
        setFileError("Укажите имя контрагента");
        return;
      }
      await rememberOrCreate(payload, { file, parsed });
    });
  };

  return (
    <Card title="Новая карточка">
      <form aria-label="Новая карточка" onSubmit={handleSubmit} style={{ display: "grid", gap: "0.75rem" }}>
        <NewCardFileZone disabled={busy || paused} onFile={acceptFile} />
        <FieldWrapper label="Имя">
          <Input value={name} autoComplete="off" onChange={(event) => setName(event.target.value)} />
        </FieldWrapper>
        <FieldWrapper label="ИНН">
          <Input value={inn} autoComplete="off" onChange={(event) => setInn(event.target.value)} />
        </FieldWrapper>
        <FieldWrapper label="Номер договора">
          <Input value={number} autoComplete="off" onChange={(event) => setNumber(event.target.value)} />
        </FieldWrapper>
        {busy && (
          <p role="status" style={{ margin: 0, color: "#475467" }}>
            Разбираю карточку…
          </p>
        )}
        {fileError && <Alert tone="error">{fileError}</Alert>}
        <div>
          <Button type="submit" disabled={!name.trim() || busy || paused}>
            Создать карточку
          </Button>
        </div>
      </form>
      <DuplicateContractCardDialog
        items={duplicates}
        busy={create.isPending}
        onEdit={editDuplicate}
        onCreateAnyway={() => void createAnyway()}
        onClose={dismissDuplicates}
      />
    </Card>
  );
};

export const ContractCardList = () => {
  const [selected, setSelected] = useState<ContractCard | null>(null);
  const [intake, setIntake] = useState<ContractCardIntake | null>(null);
  const [query, setQuery] = useState("");
  const trimmed = query.trim();
  const listQuery = trimmed.length >= MIN_SEARCH_LENGTH ? trimmed : "";

  const openCard = (card: ContractCard, nextIntake: ContractCardIntake | null = null) => {
    setIntake(nextIntake);
    setSelected(card);
  };

  const closeCard = () => {
    setSelected(null);
    setIntake(null);
  };

  return (
    <>
      <Card>
        <FieldWrapper label="Поиск">
          <Input
            type="search"
            value={query}
            placeholder="Имя, ИНН или номер"
            autoComplete="off"
            onChange={(event) => setQuery(event.target.value)}
          />
        </FieldWrapper>
      </Card>
      {selected ? (
        <OpenedContractCard
          key={selected.id}
          card={selected}
          intake={intake}
          onUpdated={setSelected}
          onDeleted={closeCard}
        />
      ) : null}
      <NewContractCardForm onCreated={openCard} />
      <Card title="Занесённые карточки">
        <ContractCardResults
          query={listQuery}
          selectedId={selected?.id ?? null}
          onSelect={(card) => openCard(card)}
        />
      </Card>
    </>
  );
};
