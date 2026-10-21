import { act, renderHook, waitFor } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { createElement, type ReactNode } from "react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import {
  archiveKeys,
  useContractCardQuery,
  useContractCardsQuery,
  useCreateContractCardMutation,
  useDownloadContractDocumentMutation,
  useDownloadContractEdoAgreementMutation,
  useParseContractCardMutation,
  usePatchContractCardMutation,
  useSaveContractRequisitesMutation,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import type { ContractCard } from "@/features/commercial-archive/types/supplyContract";
import { saveBlobAs } from "@/shared/lib/downloadFile";

vi.mock("@/features/commercial-archive/api/archiveApi", () => ({
  archiveApi: {
    searchContractCards: vi.fn(),
    getContractCard: vi.fn(),
    createContractCard: vi.fn(),
    patchContractCard: vi.fn(),
    saveContractRequisites: vi.fn(),
    parseContractCard: vi.fn(),
    downloadContractDocument: vi.fn(),
    downloadContractEdoAgreement: vi.fn(),
  },
}));

vi.mock("@/shared/lib/downloadFile", () => ({
  saveBlobAs: vi.fn(),
}));

const emptyCard = (overrides: Partial<ContractCard> = {}): ContractCard => ({
  id: 9,
  name: "Секлюц",
  imported_name: "Секлюц",
  counterparty_id: null,
  number: "1034/10/26",
  contract_date: null,
  manager_name: null,
  status: "нет",
  scan_note: null,
  has_scan: false,
  legal_form: null,
  full_name: null,
  short_name: null,
  signatory_position: null,
  signatory_name: null,
  signatory_verb: null,
  authority_basis: null,
  poa_number: null,
  poa_date: null,
  inn: null,
  kpp: null,
  ogrn: null,
  legal_address: null,
  postal_address: null,
  phone: null,
  email: null,
  bank_name: null,
  account: null,
  corr_account: null,
  bik: null,
  edo_operator: null,
  edo_id: null,
  okved: null,
  requisites_complete: false,
  ...overrides,
});

const createWrapper = () => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  const wrapper = ({ children }: { children: ReactNode }) =>
    createElement(QueryClientProvider, { client }, children);
  return { client, wrapper };
};

beforeEach(() => {
  vi.clearAllMocks();
});

describe("contract card hooks", () => {
  it("searches cards under a key separate from the KP supply contract", async () => {
    vi.mocked(archiveApi.searchContractCards).mockResolvedValue({ items: [emptyCard()], count: 1 });
    const { client, wrapper } = createWrapper();
    const { result } = renderHook(() => useContractCardsQuery("секлюц"), { wrapper });

    await waitFor(() => expect(result.current.isSuccess).toBe(true));

    expect(archiveApi.searchContractCards).toHaveBeenCalledWith("секлюц");
    expect(client.getQueryData(archiveKeys.contractCards("секлюц"))).toEqual({
      items: [emptyCard()],
      count: 1,
    });
    expect(client.getQueryData(archiveKeys.supplyContract(9))).toBeUndefined();
  });

  it("does not fetch a card when the id is missing", () => {
    const { wrapper } = createWrapper();
    const { result } = renderHook(() => useContractCardQuery(null), { wrapper });

    expect(result.current.fetchStatus).toBe("idle");
    expect(archiveApi.getContractCard).not.toHaveBeenCalled();
  });

  it("stores a created card on the contract-card key", async () => {
    const card = emptyCard();
    vi.mocked(archiveApi.createContractCard).mockResolvedValue(card);
    const { client, wrapper } = createWrapper();
    const { result } = renderHook(() => useCreateContractCardMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync({ name: "Секлюц" });
    });

    expect(archiveApi.createContractCard).toHaveBeenCalledWith({ name: "Секлюц" });
    expect(client.getQueryData(archiveKeys.contractCard(card.id))).toEqual(card);
    expect(client.getQueryData(archiveKeys.supplyContract(card.id))).toBeUndefined();
  });

  it("stores a patched card on the contract-card key", async () => {
    const card = emptyCard({ number: "1035/10/26", name: "Новое имя" });
    vi.mocked(archiveApi.patchContractCard).mockResolvedValue(card);
    const { client, wrapper } = createWrapper();
    const { result } = renderHook(() => usePatchContractCardMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync({
        cardId: 9,
        payload: { name: "Новое имя", number: "1035/10/26" },
      });
    });

    expect(client.getQueryData(archiveKeys.contractCard(9))).toEqual(card);
    expect(client.getQueryData(archiveKeys.supplyContract(9))).toBeUndefined();
  });

  it("refreshes the card after requisites without writing the KP supply-contract cache", async () => {
    vi.mocked(archiveApi.saveContractRequisites).mockResolvedValue({
      id: 9,
      counterparty_id: null,
      number: "1034/10/26",
      contract_date: "2026-10-20",
      manager_name: "ivan",
      status: "нет",
      legal_form: "ooo",
      full_name: "ООО Секлюц",
      short_name: "Секлюц",
      signatory_name: "Иванов",
      email: "a@example.com",
    });
    const { client, wrapper } = createWrapper();
    client.setQueryData(archiveKeys.contractCard(9), emptyCard());
    const invalidateSpy = vi.spyOn(client, "invalidateQueries");
    const { result } = renderHook(() => useSaveContractRequisitesMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync({
        contractId: 9,
        payload: {
          legal_form: "ooo",
          full_name: "ООО Секлюц",
          short_name: "Секлюц",
          signatory_name: "Иванов",
          signatory_verb: "действующего",
          authority_basis: "устав",
          inn: "440117229007",
          ogrn: "1024400000000",
          legal_address: "Кострома",
          email: "a@example.com",
          bank_name: "Банк",
          account: "40702810000000000001",
          corr_account: "30101810000000000001",
          bik: "044525225",
        },
      });
    });

    expect(invalidateSpy).toHaveBeenCalledWith({ queryKey: archiveKeys.contractCard(9) });
    expect(client.getQueryData(archiveKeys.supplyContract(9))).toBeUndefined();
  });

  it("parses a file by contract id and leaves the KP cache untouched", async () => {
    const file = new File(["card"], "card.pdf", { type: "application/pdf" });
    vi.mocked(archiveApi.parseContractCard).mockResolvedValue({
      fields: { full_name: "ООО Секлюц" },
      doubtful: [],
      accounts: [],
      verify_failed: false,
    });
    const { client, wrapper } = createWrapper();
    const { result } = renderHook(() => useParseContractCardMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync({ contractId: 9, file });
    });

    expect(archiveApi.parseContractCard).toHaveBeenCalledWith(9, file);
    expect(client.getQueryData(archiveKeys.supplyContract(9))).toBeUndefined();
  });

  it("downloads Word through the contract document route", async () => {
    vi.mocked(archiveApi.downloadContractDocument).mockResolvedValue({
      blob: new Blob(["docx"]),
      filename: "dogovor-9.docx",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const { wrapper } = createWrapper();
    const { result } = renderHook(() => useDownloadContractDocumentMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync({ contractId: 9, format: "docx" });
    });

    expect(archiveApi.downloadContractDocument).toHaveBeenCalledWith(9, "docx");
    expect(saveBlobAs).toHaveBeenCalledWith(expect.any(Blob), "dogovor-9.docx");
  });

  it("downloads the EDO agreement through the contract id route", async () => {
    vi.mocked(archiveApi.downloadContractEdoAgreement).mockResolvedValue({
      blob: new Blob(["pdf"]),
      filename: "soglashenie-edo-9.pdf",
      contentType: "application/pdf",
    });
    const { wrapper } = createWrapper();
    const { result } = renderHook(() => useDownloadContractEdoAgreementMutation(), { wrapper });

    await act(async () => {
      await result.current.mutateAsync({ contractId: 9, format: "pdf" });
    });

    expect(archiveApi.downloadContractEdoAgreement).toHaveBeenCalledWith(9, "pdf");
    expect(saveBlobAs).toHaveBeenCalledWith(expect.any(Blob), "soglashenie-edo-9.pdf");
  });
});
