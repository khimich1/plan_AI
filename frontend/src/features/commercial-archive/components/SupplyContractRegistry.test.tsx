import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import { SupplyContractRegistry } from "@/features/commercial-archive/components/SupplyContractRegistry";
import type { SupplyContractRegistryRow } from "@/features/commercial-archive/types/supplyContract";

vi.mock("@/features/commercial-archive/api/archiveApi", () => ({
  archiveApi: {
    listSupplyContracts: vi.fn(),
    patchSupplyContract: vi.fn(),
    replaceSupplyContract: vi.fn(),
  },
}));

const activeRow: SupplyContractRegistryRow = {
  id: 1,
  number: "1028/09/26",
  contract_date: "2026-09-25",
  counterparty_name: "РОМАШКА ООО",
  counterparty_id: 7,
  manager_name: "Иван Иванов",
  status: "нет",
  scan_note: null,
  has_scan: false,
};

const cancelledRow: SupplyContractRegistryRow = {
  ...activeRow,
  id: 2,
  number: "1027/09/26",
  counterparty_name: "ЛЮТИК ООО",
  counterparty_id: 8,
  status: "отмена не будем работать",
};

const wrap = (ui: ReactNode) => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  return render(<QueryClientProvider client={client}>{ui}</QueryClientProvider>);
};

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe("SupplyContractRegistry", () => {
  it("patches status and does not replace when choosing cancel", async () => {
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([activeRow]);
    vi.mocked(archiveApi.patchSupplyContract).mockResolvedValue({
      id: 1,
      counterparty_id: 7,
      number: "1028/09/26",
      contract_date: "2026-09-25",
      manager_name: "Иван Иванов",
      status: "отмена не будем работать",
      legal_form: "ooo",
      full_name: "Ромашка",
      short_name: "Ромашка",
      signatory_name: "Иванов",
      email: "a@b.ru",
    });

    wrap(<SupplyContractRegistry open onRequestNew={vi.fn()} />);

    expect(await screen.findByText("1028/09/26")).toBeInTheDocument();
    expect(screen.getByText("РОМАШКА ООО")).toBeInTheDocument();
    expect(screen.getByText("Иван Иванов")).toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Статус 1028/09/26"), {
      target: { value: "отмена не будем работать" },
    });

    await waitFor(() => {
      expect(archiveApi.patchSupplyContract).toHaveBeenCalledWith(1, {
        status: "отмена не будем работать",
      });
    });
    expect(archiveApi.replaceSupplyContract).not.toHaveBeenCalled();
  });

  it("shows Новый договор only on a cancelled row", async () => {
    vi.mocked(archiveApi.listSupplyContracts).mockResolvedValue([activeRow, cancelledRow]);
    const onRequestNew = vi.fn();

    wrap(<SupplyContractRegistry open onRequestNew={onRequestNew} />);

    const active = (await screen.findByText("1028/09/26")).closest("tr");
    const cancelled = screen.getByText("1027/09/26").closest("tr");
    expect(active).toBeTruthy();
    expect(cancelled).toBeTruthy();
    expect(within(active as HTMLElement).queryByRole("button", { name: "Новый договор" })).not.toBeInTheDocument();
    fireEvent.click(within(cancelled as HTMLElement).getByRole("button", { name: "Новый договор" }));
    expect(onRequestNew).toHaveBeenCalledWith(expect.objectContaining({ id: 2, number: "1027/09/26" }));
    expect(archiveApi.replaceSupplyContract).not.toHaveBeenCalled();
  });
});
