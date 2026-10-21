import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Mock } from "vitest";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import { httpClient } from "@/shared/api/httpClient";

vi.mock("@/shared/api/httpClient", () => ({
  httpClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    patch: vi.fn(),
    delete: vi.fn(),
    download: vi.fn(),
  },
  resolveApiUrl: (path: string) => path,
}));

const mockGet = httpClient.get as unknown as Mock;
const mockPost = httpClient.post as unknown as Mock;
const mockPatch = httpClient.patch as unknown as Mock;
const mockDelete = httpClient.delete as unknown as Mock;
const mockDownload = httpClient.download as unknown as Mock;

const contractCardQuery = (path: string): string | null => {
  const url = new URL(path, "http://localhost");
  return url.pathname === "/api/v1/commercial/archive/contract-cards"
    ? url.searchParams.get("q")
    : null;
};

beforeEach(() => {
  vi.clearAllMocks();
});

describe("archiveApi MNA-602 — resume", () => {
  it("exposes resume helper", () => {
    expect(typeof (archiveApi as { resume?: unknown }).resume).toBe("function");
  });

  it("POSTs /api/v1/commercial/archive/{kpId}/resume", async () => {
    mockPost.mockResolvedValue({ draft_id: "draft-resume-42" });

    const resume = (archiveApi as { resume: (kpId: number) => Promise<unknown> }).resume;
    await resume(42);

    expect(mockPost).toHaveBeenCalledWith("/api/v1/commercial/archive/42/resume");
  });
});

describe("archiveApi xlsx_delivery_in_unit", () => {
  it("builds document URL for the embed kind", () => {
    expect(archiveApi.buildDocumentUrl(7, "xlsx_delivery_in_unit")).toBe(
      "/api/v1/commercial/archive/7/files/xlsx_delivery_in_unit",
    );
  });

  it("uses suffix fallback filename for embed xlsx", async () => {
    mockDownload.mockResolvedValue(undefined);
    await archiveApi.downloadDocument(7, "xlsx_delivery_in_unit");
    expect(mockDownload).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/7/files/xlsx_delivery_in_unit",
      "КП_7_с_доставкой_в_цене.xlsx",
    );
  });

  it("keeps ordinary xlsx fallback name", async () => {
    mockDownload.mockResolvedValue(undefined);
    await archiveApi.downloadDocument(7, "xlsx");
    expect(mockDownload).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/7/files/xlsx",
      "КП_7.xlsx",
    );
  });
});

describe("archiveApi bindCounterparty", () => {
  it("PATCHes /counterparty with the chosen id", async () => {
    mockPatch.mockResolvedValue({ kp_id: 42, counterparty_id: 7 });

    await archiveApi.bindCounterparty(42, 7);

    expect(mockPatch).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/42/counterparty",
      JSON.stringify({ counterparty_id: 7 }),
      { "Content-Type": "application/json" },
    );
  });
});

describe("archiveApi createAndBindCounterparty", () => {
  it("POSTs /counterparty with name and code_1c", async () => {
    mockPost.mockResolvedValue({ kp_id: 42, counterparty_id: 11 });

    await archiveApi.createAndBindCounterparty(42, {
      name: "ООО Бармалей",
      code_1c: "00-BARM",
      inn: "7701000001",
      kpp: "770101001",
    });

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/42/counterparty",
      JSON.stringify({
        name: "ООО Бармалей",
        code_1c: "00-BARM",
        inn: "7701000001",
        kpp: "770101001",
      }),
      { "Content-Type": "application/json" },
    );
    expect(mockPatch).not.toHaveBeenCalled();
  });
});

describe("archiveApi contract cards", () => {
  it("GETs /contract-cards with the search text", async () => {
    mockGet.mockResolvedValue({ items: [], count: 0 });

    await archiveApi.searchContractCards("секлюц");

    expect(mockGet).toHaveBeenCalledTimes(1);
    const path = mockGet.mock.calls[0][0] as string;
    expect(path.startsWith("/api/v1/commercial/archive/contract-cards?")).toBe(true);
    expect(contractCardQuery(path)).toBe("секлюц");
  });

  it("GETs /contract-cards/duplicates with name and inn", async () => {
    mockGet.mockResolvedValue({ items: [], count: 0 });

    await archiveApi.findDuplicateContractCards("ооо  ромашка", "5024 243403");

    expect(mockGet).toHaveBeenCalledTimes(1);
    const path = mockGet.mock.calls[0][0] as string;
    const url = new URL(path, "http://localhost");
    expect(url.pathname).toBe("/api/v1/commercial/archive/contract-cards/duplicates");
    expect(url.searchParams.get("name")).toBe("ооо  ромашка");
    expect(url.searchParams.get("inn")).toBe("5024 243403");
  });

  it("POSTs a card with only the name", async () => {
    mockPost.mockResolvedValue({ id: 1, name: "Секлюц" });

    await archiveApi.createContractCard({ name: "Секлюц" });

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/contract-cards",
      JSON.stringify({ name: "Секлюц" }),
      { "Content-Type": "application/json" },
    );
  });

  it("GETs one card by id", async () => {
    mockGet.mockResolvedValue({ id: 9, name: "Секлюц" });

    await archiveApi.getContractCard(9);

    expect(mockGet).toHaveBeenCalledWith("/api/v1/commercial/archive/contract-cards/9");
  });

  it("PATCHes number, name and inn on the card", async () => {
    mockPatch.mockResolvedValue({ id: 9, number: "1035/10/26" });

    await archiveApi.patchContractCard(9, {
      name: "Новое имя",
      inn: "440117229007",
      number: "1035/10/26",
    });

    expect(mockPatch).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/contract-cards/9",
      JSON.stringify({
        name: "Новое имя",
        inn: "440117229007",
        number: "1035/10/26",
      }),
      { "Content-Type": "application/json" },
    );
  });

  it("DELETEs one card by id without a confirmation word", async () => {
    mockDelete.mockResolvedValue(undefined);

    await archiveApi.deleteContractCard(9);

    expect(mockDelete).toHaveBeenCalledTimes(1);
    expect(mockDelete).toHaveBeenCalledWith("/api/v1/commercial/archive/contract-cards/9");
  });

  it("POSTs requisites for the contract id", async () => {
    const payload = {
      legal_form: "ooo" as const,
      full_name: "ООО Секлюц",
      short_name: "Секлюц",
      signatory_name: "Иванов Иван",
      signatory_verb: "действующего" as const,
      authority_basis: "устав" as const,
      inn: "440117229007",
      ogrn: "1024400000000",
      legal_address: "Кострома",
      email: "a@example.com",
      bank_name: "Банк",
      account: "40702810000000000001",
      corr_account: "30101810000000000001",
      bik: "044525225",
    };
    mockPost.mockResolvedValue({ id: 9, number: "1034/10/26" });

    await archiveApi.saveContractRequisites(9, payload);

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/supply-contracts/9/requisites",
      JSON.stringify(payload),
      { "Content-Type": "application/json" },
    );
  });

  it("POSTs a new card file to parse without a contract id", async () => {
    const file = new File(["card"], "card.docx");
    mockPost.mockResolvedValue({ fields: {}, doubtful: [], accounts: [], verify_failed: false });

    await archiveApi.parseContractCardFile(file);

    expect(mockPost).toHaveBeenCalledTimes(1);
    const [path, body] = mockPost.mock.calls[0] as [string, FormData];
    expect(path).toBe("/api/v1/commercial/archive/contract-cards/parse");
    expect(path).not.toContain("/supply-contracts/");
    expect(path).not.toContain("/import");
    expect(body).toBeInstanceOf(FormData);
    expect(body.get("file")).toBe(file);
  });

  it("POSTs the card file to parse by contract id", async () => {
    const file = new File(["card"], "card.pdf", { type: "application/pdf" });
    mockPost.mockResolvedValue({ fields: {}, doubtful: [], accounts: [], verify_failed: false });

    await archiveApi.parseContractCard(9, file);

    expect(mockPost).toHaveBeenCalledTimes(1);
    const [path, body] = mockPost.mock.calls[0] as [string, FormData];
    expect(path).toBe("/api/v1/commercial/archive/supply-contracts/9/parse");
    expect(body).toBeInstanceOf(FormData);
    expect(body.get("file")).toBe(file);
  });

  it("downloads the document by contract id, not by kp id", async () => {
    mockDownload.mockResolvedValue({ blob: new Blob(), filename: "dogovor-15.docx" });

    await archiveApi.downloadContractDocument(15, "docx");

    const path = mockDownload.mock.calls[0][0] as string;
    expect(path).toBe("/api/v1/commercial/archive/supply-contracts/15/document?format=docx");
    expect(path).not.toContain("/supply-contract/document");
  });

  it("downloads the document PDF with format=pdf on the contract id route", async () => {
    mockDownload.mockResolvedValue({ blob: new Blob(), filename: "dogovor-15.pdf" });

    await archiveApi.downloadContractDocument(15, "pdf");

    expect(mockDownload).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/supply-contracts/15/document?format=pdf",
      "dogovor-15.pdf",
    );
  });

  it("downloads the EDO agreement by contract id", async () => {
    mockDownload.mockResolvedValue({ blob: new Blob(), filename: "soglashenie-edo-15.pdf" });

    await archiveApi.downloadContractEdoAgreement(15, "pdf");

    expect(mockDownload).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/supply-contracts/15/edo-agreement?format=pdf",
      "soglashenie-edo-15.pdf",
    );
  });

  it("keeps the KP document download on the kp route", async () => {
    mockDownload.mockResolvedValue({ blob: new Blob(), filename: "dogovor-42.pdf" });

    await archiveApi.downloadSupplyContract(42, "pdf");

    expect(mockDownload).toHaveBeenCalledWith(
      "/api/v1/commercial/archive/42/supply-contract/document?format=pdf",
      "dogovor-42.pdf",
    );
  });
});
