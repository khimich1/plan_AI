import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { Mock } from "vitest";
import { MemoryRouter, useLocation } from "react-router";
import { WizardDraftProvider } from "@/features/commercial-offer/store/wizardDraftStore";
import { httpClient } from "@/shared/api/httpClient";
import {
  CommercialOfferArchivePage,
  concreteProductTypes,
  filterByProductType,
  sectionFromStatus,
} from "@/pages/commercial-offer-archive/CommercialOfferArchivePage";
import type { ArchiveOfferListItem } from "@/features/commercial-archive/types/archive";

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

const SearchProbe = () => {
  const { search } = useLocation();
  return <output aria-label="Адрес раздела">{search}</output>;
};

const renderArchivePage = (initialEntry: string) => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false } },
  });
  return render(
    <QueryClientProvider client={client}>
      <WizardDraftProvider>
        <MemoryRouter initialEntries={[initialEntry]}>
          <SearchProbe />
          <CommercialOfferArchivePage />
        </MemoryRouter>
      </WizardDraftProvider>
    </QueryClientProvider>,
  );
};

const archiveListCalls = (): string[] =>
  mockGet.mock.calls
    .map(([url]) => String(url))
    .filter((url) => url.startsWith("/api/v1/commercial/archive?"));

const baseItem: ArchiveOfferListItem = {
  kp_id: 1,
  creation_date: "01.03.2026",
  customer_name: "Клиент",
  manager_name: "Менеджер",
  discount_percent: 0,
  subtotal: 1000,
  vat_amount: 220,
  total_amount: 1220,
  execution_terms: null,
  status: "в архиве",
  completion_percentage: null,
};

describe("sectionFromStatus", () => {
  it("maps На СГП to in_production", () => {
    expect(sectionFromStatus({ ...baseItem, status: "На СГП" })).toBe("in_production");
  });

  it("maps в работе to in_production", () => {
    expect(sectionFromStatus({ ...baseItem, status: "в работе" })).toBe("in_production");
  });

  it("maps выполнено to completed", () => {
    expect(sectionFromStatus({ ...baseItem, status: "выполнено" })).toBe("completed");
  });

  it("maps в архиве to archived", () => {
    expect(sectionFromStatus({ ...baseItem, status: "в архиве" })).toBe("archived");
  });

  it("maps на согласовании to on_approval", () => {
    expect(sectionFromStatus({ ...baseItem, status: "на согласовании" })).toBe("on_approval");
  });
});

describe("filterByProductType MNA-602 — contains-type", () => {
  const mixedWithPlates: ArchiveOfferListItem = {
    ...baseItem,
    kp_id: 10,
    product_type: "mixed",
    product_types: ["plates", "piles"],
  };

  const monoPiles: ArchiveOfferListItem = {
    ...baseItem,
    kp_id: 11,
    product_type: "piles",
  };

  const monoPlates: ArchiveOfferListItem = {
    ...baseItem,
    kp_id: 12,
    product_type: "plates",
  };

  it("keeps mixed KP when product_types contains the filter", () => {
    const result = filterByProductType(
      [mixedWithPlates, monoPiles, monoPlates],
      "plates",
    );
    expect(result.map((item) => item.kp_id)).toEqual([10, 12]);
  });

  it("matches piles filter via product_types on mixed KP", () => {
    const result = filterByProductType([mixedWithPlates, monoPlates], "piles");
    expect(result.map((item) => item.kp_id)).toEqual([10]);
  });

  it("ignores literal mixed token in product_types", () => {
    expect(
      concreteProductTypes({
        ...baseItem,
        product_type: "mixed",
        product_types: ["mixed", "steps"],
      }),
    ).toEqual(["steps"]);
  });

  it("returns all items when filter is all", () => {
    const items = [mixedWithPlates, monoPiles];
    expect(filterByProductType(items, "all")).toEqual(items);
  });
});

describe("CommercialOfferArchivePage counterparties section", () => {
  afterEach(() => {
    cleanup();
    mockGet.mockReset();
  });

  it("writes section=counterparties and does not request the KP list", async () => {
    mockGet.mockImplementation((url: string) =>
      String(url).includes("contract-cards")
        ? Promise.resolve({ items: [], count: 0 })
        : Promise.resolve([baseItem]),
    );
    renderArchivePage("/archive?section=archived");

    expect(await screen.findByText("КП №1")).toBeInTheDocument();

    fireEvent.click(screen.getByRole("tab", { name: "Контрагенты" }));

    expect(screen.getByLabelText("Адрес раздела")).toHaveTextContent("section=counterparties");
    expect(screen.getByRole("tab", { name: "Контрагенты" })).toHaveAttribute("aria-selected", "true");
    expect(await screen.findByText("Карточек пока нет")).toBeInTheDocument();
    expect(screen.getByRole("searchbox", { name: "Поиск" })).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Создать карточку" })).toBeDisabled();
    expect(screen.queryByRole("tablist", { name: "Тип продукции" })).not.toBeInTheDocument();
    expect(screen.queryByPlaceholderText("Номер КП")).not.toBeInTheDocument();
    expect(screen.queryByPlaceholderText("Название заказчика")).not.toBeInTheDocument();
    expect(screen.queryByText("КП №1")).not.toBeInTheDocument();
    expect(archiveListCalls().some((url) => url.includes("section=counterparties"))).toBe(false);
    expect(mockGet.mock.calls.map(([url]) => String(url)).some((url) => url.includes("contract-cards"))).toBe(
      true,
    );
  });

  it("opens counterparties from the query without requesting the KP list", async () => {
    mockGet.mockImplementation((url: string) =>
      String(url).includes("contract-cards")
        ? Promise.resolve({ items: [], count: 0 })
        : Promise.resolve([baseItem]),
    );
    renderArchivePage("/archive?section=counterparties");

    expect(await screen.findByText("Карточек пока нет")).toBeInTheDocument();
    expect(screen.getByRole("searchbox", { name: "Поиск" })).toBeInTheDocument();
    expect(screen.getByLabelText("Имя")).toBeInTheDocument();
    expect(screen.getByRole("tab", { name: "Контрагенты" })).toHaveAttribute("aria-selected", "true");
    expect(screen.queryByRole("tablist", { name: "Тип продукции" })).not.toBeInTheDocument();
    expect(screen.queryByPlaceholderText("Номер КП")).not.toBeInTheDocument();
    expect(archiveListCalls()).toEqual([]);
    expect(mockGet.mock.calls.map(([url]) => String(url)).some((url) => url.includes("contract-cards"))).toBe(
      true,
    );
  });

  it("still shows the KP list on the archive section", async () => {
    mockGet.mockResolvedValue([{ ...baseItem, kp_id: 42 }]);
    renderArchivePage("/archive?section=archived");

    expect(await screen.findByText("КП №42")).toBeInTheDocument();
    expect(mockGet).toHaveBeenCalledWith("/api/v1/commercial/archive?section=archived");
    expect(screen.queryByText("Введите имя, ИНН или номер.")).not.toBeInTheDocument();
    expect(screen.queryByRole("searchbox", { name: "Поиск" })).not.toBeInTheDocument();
    expect(screen.getByRole("tablist", { name: "Тип продукции" })).toBeInTheDocument();
    expect(screen.getByPlaceholderText("Номер КП")).toBeInTheDocument();
  });
});
