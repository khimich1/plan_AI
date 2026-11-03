import { useEffect, useRef, useState, type CSSProperties } from "react";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import {
  useDirectoryQuery,
  useEnsureDirectoryContractMutation,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import type {
  ContractCard,
  DirectoryClient,
  DirectoryContractLink,
} from "@/features/commercial-archive/types/supplyContract";
import { ContractCardDrawer } from "./ContractCardDrawer";
import { getErrorMessage } from "@/shared/lib/apiError";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { FieldWrapper, Input } from "@/shared/ui/Field";

const MIN_SEARCH_LENGTH = 2;
const COLUMN_GRID = "minmax(0, 1.6fr) minmax(8.5rem, 0.9fr) minmax(10rem, auto)";

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
  textAlign: "left",
  font: "inherit",
  color: "inherit",
});

const OpenedContractCard = ({
  card,
  onUpdated,
  onDeleted,
}: {
  card: ContractCard;
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
        <ContractCardDrawer card={card} onUpdated={onUpdated} onDeleted={onDeleted} />
      </Card>
    </section>
  );
};

export const DirectoryContractActions = ({
  client,
  pending,
  error,
  onLink,
  onAssemble,
}: {
  client: DirectoryClient;
  pending: boolean;
  error: string | null;
  onLink: (body: DirectoryContractLink) => void;
  onAssemble?: () => void;
}) => {
  const [typed, setTyped] = useState("");
  const linkable = client.linkable;
  const submitTyped = () => {
    const number = typed.trim();
    if (!number || pending) {
      return;
    }
    onLink({ number });
  };

  return (
    <span style={{ display: "flex", gap: "0.75rem", alignItems: "center", flexWrap: "wrap" }}>
      <span style={{ color: "#b54708" }}>договор нет</span>
      {linkable ? (
        <Button
          type="button"
          variant="secondary"
          disabled={pending}
          onClick={() => onLink({ contract_id: linkable.id })}
        >
          {`Привязать № ${linkable.number}`}
        </Button>
      ) : (
        <>
          {client.candidates.map((candidate) => (
            <Button
              key={candidate.id}
              type="button"
              variant="secondary"
              disabled={pending}
              onClick={() => onLink({ contract_id: candidate.id })}
            >
              {`Привязать № ${candidate.number}`}
            </Button>
          ))}
          <form
            aria-label="Старый номер"
            onSubmit={(event) => {
              event.preventDefault();
              submitTyped();
            }}
            style={{ display: "flex", gap: "0.5rem", alignItems: "end", flexWrap: "wrap" }}
          >
            <FieldWrapper label="Номер договора">
              <Input
                value={typed}
                autoComplete="off"
                onChange={(event) => setTyped(event.target.value)}
              />
            </FieldWrapper>
            <Button type="submit" variant="secondary" disabled={pending || typed.trim() === ""}>
              Привязать
            </Button>
          </form>
        </>
      )}
      {onAssemble ? (
        <Button type="button" variant="secondary" disabled={pending} onClick={onAssemble}>
          Собрать договор
        </Button>
      ) : null}
      {error ? <Alert tone="error">{error}</Alert> : null}
    </span>
  );
};

const DirectoryRow = ({
  item,
  selected,
  pending,
  error,
  onOpen,
  onLink,
  onAssemble,
}: {
  item: DirectoryClient;
  selected: boolean;
  pending: boolean;
  error: string | null;
  onOpen: (contractId: number) => void;
  onLink: (body: DirectoryContractLink) => void;
  onAssemble?: () => void;
}) => {
  const inn = listInn(item.inn);
  const contract = item.contract;
  if (contract) {
    return (
      <button
        type="button"
        aria-pressed={selected}
        onClick={() => onOpen(contract.id)}
        style={{ ...rowStyle(selected), cursor: "pointer" }}
      >
        <span style={{ fontWeight: 600, overflowWrap: "anywhere" }}>{item.name}</span>
        <span>{inn}</span>
        <span>{contract.number}</span>
      </button>
    );
  }

  return (
    <div style={rowStyle(false)}>
      <span style={{ fontWeight: 600, overflowWrap: "anywhere" }}>{item.name}</span>
      <span>{inn}</span>
      <DirectoryContractActions
        client={item}
        pending={pending}
        error={error}
        onLink={onLink}
        onAssemble={onAssemble}
      />
    </div>
  );
};

const DirectoryResults = ({
  query,
  selectedContractId,
  onOpen,
  onAssemble,
}: {
  query: string;
  selectedContractId: number | null;
  onOpen: (contractId: number) => void;
  onAssemble: (client: DirectoryClient) => void;
}) => {
  const search = useDirectoryQuery(query);
  const ensure = useEnsureDirectoryContractMutation();
  const [rowError, setRowError] = useState<{ id: number; message: string } | null>(null);

  const linkContract = (item: DirectoryClient, body: DirectoryContractLink) => {
    setRowError(null);
    void ensure
      .mutateAsync({
        counterpartyId: item.id,
        ...(body.number ? { number: body.number } : {}),
        ...(body.contract_id != null ? { contractId: body.contract_id } : {}),
      })
      .catch((error: unknown) => setRowError({ id: item.id, message: getErrorMessage(error) }));
  };

  if (search.isError) {
    return <Alert tone="error">{getErrorMessage(search.error)}</Alert>;
  }
  if (!search.data) {
    return (
      <p role="status" style={{ margin: 0, color: "#475467" }}>
        Ищу контрагентов…
      </p>
    );
  }
  if (search.data.items.length === 0) {
    return (
      <p role="status" style={{ margin: 0, color: "#475467" }}>
        Ничего не найдено
      </p>
    );
  }

  return (
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
        <span>Договор</span>
      </div>
      <ul
        aria-label="Контрагенты"
        style={{ listStyle: "none", margin: 0, padding: 0, display: "grid", gap: "0.5rem" }}
      >
        {search.data.items.map((item) => (
          <li key={item.id}>
            <DirectoryRow
              item={item}
              selected={item.contract?.id === selectedContractId}
              pending={ensure.isPending && ensure.variables?.counterpartyId === item.id}
              error={rowError?.id === item.id ? rowError.message : null}
              onOpen={onOpen}
              onLink={(body) => linkContract(item, body)}
              onAssemble={() => onAssemble(item)}
            />
          </li>
        ))}
      </ul>
    </div>
  );
};

export const ContractCardList = () => {
  const [selected, setSelected] = useState<ContractCard | null>(null);
  const [draftClient, setDraftClient] = useState<DirectoryClient | null>(null);
  const [query, setQuery] = useState("");
  const [openError, setOpenError] = useState<string | null>(null);
  const trimmed = query.trim();
  const listQuery = trimmed.length >= MIN_SEARCH_LENGTH ? trimmed : "";

  const openContract = async (contractId: number) => {
    setDraftClient(null);
    setOpenError(null);
    try {
      const card = await archiveApi.getContractCard(contractId);
      setSelected(card);
    } catch (error) {
      setOpenError(getErrorMessage(error));
    }
  };

  const closeCard = () => {
    setSelected(null);
  };

  return (
    <>
      <Card>
        <FieldWrapper label="Поиск">
          <Input
            type="search"
            value={query}
            placeholder="Имя или ИНН"
            autoComplete="off"
            onChange={(event) => setQuery(event.target.value)}
          />
        </FieldWrapper>
      </Card>
      {selected ? (
        <OpenedContractCard
          key={selected.id}
          card={selected}
          onUpdated={setSelected}
          onDeleted={closeCard}
        />
      ) : null}
      {draftClient ? (
        <Card title={draftClient.name}>
          <ContractCardDrawer
            key={draftClient.id}
            draft={{
              counterpartyId: draftClient.id,
              name: draftClient.name,
              inn: draftClient.inn,
            }}
            onCloseDraft={() => setDraftClient(null)}
          />
        </Card>
      ) : null}
      {openError && <Alert tone="error">{openError}</Alert>}
      <Card title="Контрагенты">
        {listQuery ? (
          <DirectoryResults
            query={listQuery}
            selectedContractId={selected?.id ?? null}
            onOpen={(contractId) => void openContract(contractId)}
            onAssemble={(client) => {
              setSelected(null);
              setDraftClient(client);
            }}
          />
        ) : (
          <p role="status" style={{ margin: 0, color: "#475467" }}>
            Введите имя или ИНН
          </p>
        )}
      </Card>
    </>
  );
};
