import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Mock } from "vitest";
import { logisticsApi } from "@/features/logistics/api/logisticsApi";
import { httpClient } from "@/shared/api/httpClient";

vi.mock("@/shared/api/httpClient", () => ({
  httpClient: {
    get: vi.fn(),
    post: vi.fn(),
    put: vi.fn(),
    patch: vi.fn(),
    download: vi.fn(),
  },
}));

const mockGet = httpClient.get as unknown as Mock;
const mockPost = httpClient.post as unknown as Mock;
const mockPut = httpClient.put as unknown as Mock;

const JSON_HEADERS = { "Content-Type": "application/json" };

beforeEach(() => {
  vi.clearAllMocks();
});

// Phase 6 contract regression: backend list endpoints return an {items, count}
// envelope — the api layer must unwrap it (the UI previously mapped over the
// raw envelope object and crashed at runtime).
describe("logisticsApi envelope {items, count} unwrapping", () => {
  it("listShipments resolves to the rows array from the envelope", async () => {
    const rows = [{ id: 1 }, { id: 2 }];
    mockGet.mockResolvedValue({ items: rows, count: 2 });

    const result = await logisticsApi.listShipments();

    expect(mockGet).toHaveBeenCalledWith("/api/v1/logistics/shipments");
    expect(result).toEqual(rows);
  });

  it("listCarriers resolves to the carriers array from the envelope", async () => {
    const carriers = [{ id: 1, name: "ООО АвтоЛайн" }];
    mockGet.mockResolvedValue({ items: carriers, count: 1 });

    const result = await logisticsApi.listCarriers({ q: "авто", activeOnly: true });

    expect(mockGet).toHaveBeenCalledWith("/api/v1/logistics/carriers?q=%D0%B0%D0%B2%D1%82%D0%BE&active=1");
    expect(result).toEqual(carriers);
  });

  it("searchPileCatalog resolves to the entries array from the envelope", async () => {
    const entries = [{ id: 1, mark: "С60.30", weight_kg: 1380 }];
    mockGet.mockResolvedValue({ items: entries, count: 1 });

    const result = await logisticsApi.searchPileCatalog("С60");

    expect(mockGet).toHaveBeenCalledWith(
      `/api/v1/logistics/pile-catalog?q=${encodeURIComponent("С60")}`,
    );
    expect(result).toEqual(entries);
  });

  it("searchKp calls /kp-search with kp_id or customer", async () => {
    const response = {
      mode: "number",
      items: [{ kp_id: 10, customer_name: "А", status: "в работе", product_type: "plates" }],
      total: 1,
      truncated: false,
    };
    mockGet.mockResolvedValue(response);

    await expect(logisticsApi.searchKp({ kpId: 10 })).resolves.toEqual(response);
    expect(mockGet).toHaveBeenCalledWith("/api/v1/logistics/kp-search?kp_id=10");

    mockGet.mockResolvedValue({ ...response, mode: "customer" });
    await logisticsApi.searchKp({ customer: "Ромашка" });
    expect(mockGet).toHaveBeenCalledWith(
      `/api/v1/logistics/kp-search?customer=${encodeURIComponent("Ромашка")}`,
    );
  });
});

// Phase 6 contract regression: backend reads vehicle_class from the query
// string — a JSON body would be silently ignored.
describe("logisticsApi proposeItems vehicle_class", () => {
  it("sends vehicle_class as a query param, not a body", async () => {
    mockPost.mockResolvedValue({ items: [], not_fit: [] });

    await logisticsApi.proposeItems(5, "t20");

    expect(mockPost).toHaveBeenCalledWith("/api/v1/logistics/shipments/5/propose?vehicle_class=t20");
  });

  it("requests the bare propose path when vehicle class is not set", async () => {
    mockPost.mockResolvedValue({ items: [], not_fit: [] });

    await logisticsApi.proposeItems(5);

    expect(mockPost).toHaveBeenCalledWith("/api/v1/logistics/shipments/5/propose");
  });
});

// Phase 6 contract regression: PUT /items returns the full shipment card,
// complete/cancel return {ok, shipment_id, status, message} — not ShipmentItem[]
// and not ShipmentDetails.
describe("logisticsApi reuseTransport", () => {
  it("POSTs create payload to /shipments/{sourceId}/reuse-transport", async () => {
    const card = { id: 9, status: "in_work", carrier_id: 3 };
    mockPost.mockResolvedValue(card);
    const payload = {
      shipment_date: "2026-08-02",
      delivery_type: "pickup" as const,
      kp_ids: [160],
    };

    const result = await logisticsApi.reuseTransport(5, payload);

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/logistics/shipments/5/reuse-transport",
      JSON.stringify(payload),
      JSON_HEADERS,
    );
    expect(result).toBe(card);
  });
});

describe("logisticsApi mutation response contracts", () => {
  it("confirmItems PUTs the {items} payload and returns the shipment card as-is", async () => {
    const card = { id: 5, status: "in_work", items: [{ id: 9, qty: 3 }] };
    mockPut.mockResolvedValue(card);
    const items = [{ item_type: "plate" as const, completed_plate_id: 77, qty: 3, sort_order: 0 }];

    const result = await logisticsApi.confirmItems(5, items);

    expect(mockPut).toHaveBeenCalledWith(
      "/api/v1/logistics/shipments/5/items",
      JSON.stringify({ items }),
      JSON_HEADERS,
    );
    expect(result).toBe(card);
  });

  it("completeShipment returns the mutation result envelope as-is", async () => {
    const mutation = { ok: true, shipment_id: 5, status: "done", message: "Рейс #5 обработан" };
    mockPost.mockResolvedValue(mutation);

    const result = await logisticsApi.completeShipment(5);

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/logistics/shipments/5/complete",
      JSON.stringify({}),
      JSON_HEADERS,
    );
    expect(result).toBe(mutation);
  });

  it("cancelShipment returns the mutation result envelope as-is", async () => {
    const mutation = { ok: true, shipment_id: 5, status: "cancelled", message: "Рейс #5 отменён" };
    mockPost.mockResolvedValue(mutation);

    const result = await logisticsApi.cancelShipment(5);

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/logistics/shipments/5/cancel",
      JSON.stringify({}),
      JSON_HEADERS,
    );
    expect(result).toBe(mutation);
  });
});

describe("logisticsApi listShipments filter query", () => {
  it("serializes every registry filter into the query string", async () => {
    mockGet.mockResolvedValue({ items: [], count: 0 });

    await logisticsApi.listShipments({
      date_from: "2026-08-01",
      date_to: "2026-08-31",
      kp_id: 154,
      carrier_id: 3,
      delivery_type: "pickup",
      status: "done",
      no_upd: true,
      attention: true,
    });

    const url = mockGet.mock.calls[0][0] as string;
    expect(url).toContain("date_from=2026-08-01");
    expect(url).toContain("date_to=2026-08-31");
    expect(url).toContain("kp_id=154");
    expect(url).toContain("carrier_id=3");
    expect(url).toContain("delivery_type=pickup");
    expect(url).toContain("status=done");
    expect(url).toContain("no_upd=1");
    expect(url).toContain("attention=1");
  });

  it("omits unset filters from the query string", async () => {
    mockGet.mockResolvedValue({ items: [], count: 0 });

    await logisticsApi.listShipments({ attention: true });

    const url = mockGet.mock.calls[0][0] as string;
    expect(url).toBe("/api/v1/logistics/shipments?attention=1");
  });
});
