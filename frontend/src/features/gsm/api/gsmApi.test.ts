import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Mock } from "vitest";
import { gsmApi } from "@/features/gsm/api/gsmApi";
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
const mockPatch = httpClient.patch as unknown as Mock;
const mockDownload = httpClient.download as unknown as Mock;

const JSON_HEADERS = { "Content-Type": "application/json" };

beforeEach(() => {
  vi.clearAllMocks();
});

describe("gsmApi registry list contracts", () => {
  it("listVehicles hits /gsm/vehicles and returns the array as-is", async () => {
    const vehicles = [{ id: 1, name: "Geely", plate_number: "A123" }];
    mockGet.mockResolvedValue(vehicles);

    await expect(gsmApi.listVehicles()).resolves.toEqual(vehicles);
    expect(mockGet).toHaveBeenCalledWith("/api/v1/gsm/vehicles");
  });

  it("listVehicles passes active_only=false when requested", async () => {
    mockGet.mockResolvedValue([]);
    await gsmApi.listVehicles({ activeOnly: false });
    expect(mockGet).toHaveBeenCalledWith("/api/v1/gsm/vehicles?active_only=false");
  });

  it("listDrivers hits /gsm/drivers", async () => {
    mockGet.mockResolvedValue([]);
    await gsmApi.listDrivers();
    expect(mockGet).toHaveBeenCalledWith("/api/v1/gsm/drivers");
  });

  it("listCards hits /gsm/cards and can include archived", async () => {
    mockGet.mockResolvedValue([]);
    await gsmApi.listCards();
    expect(mockGet).toHaveBeenCalledWith("/api/v1/gsm/cards");

    await gsmApi.listCards({ includeArchived: true });
    expect(mockGet).toHaveBeenCalledWith("/api/v1/gsm/cards?include_archived=true");
  });

  it("listStations hits /gsm/stations", async () => {
    mockGet.mockResolvedValue([]);
    await gsmApi.listStations();
    expect(mockGet).toHaveBeenCalledWith("/api/v1/gsm/stations");
  });

  it("getSettings hits /gsm/settings", async () => {
    const settings = { winter_start: "11-01", hook_threshold_km: 13 };
    mockGet.mockResolvedValue(settings);
    await expect(gsmApi.getSettings()).resolves.toEqual(settings);
    expect(mockGet).toHaveBeenCalledWith("/api/v1/gsm/settings");
  });
});

describe("gsmApi mutations", () => {
  it("createVehicle POSTs JSON payload", async () => {
    const payload = {
      name: "Geely",
      plate_number: "A123BC",
      tank_volume_liters: 60,
      norm_summer: 12,
      norm_winter: 14,
    };
    mockPost.mockResolvedValue({ id: 1, ...payload, primary_driver_id: null, is_active: true });

    await gsmApi.createVehicle(payload);

    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/gsm/vehicles",
      JSON.stringify(payload),
      JSON_HEADERS,
    );
  });

  it("patchCard PATCHes archive / vehicle bind", async () => {
    mockPatch.mockResolvedValue({ id: 3, card_number: "100", vehicle_id: 2, assigned_at: "2026-01-01", archived_at: null });
    await gsmApi.patchCard(3, { vehicle_id: 2, archive: false });
    expect(mockPatch).toHaveBeenCalledWith(
      "/api/v1/gsm/cards/3",
      JSON.stringify({ vehicle_id: 2, archive: false }),
      JSON_HEADERS,
    );
  });

  it("putSettings PUTs settings body", async () => {
    const payload = { winter_start: "11-15", hook_threshold_km: 10 };
    mockPut.mockResolvedValue(payload);
    await gsmApi.putSettings(payload);
    expect(mockPut).toHaveBeenCalledWith(
      "/api/v1/gsm/settings",
      JSON.stringify(payload),
      JSON_HEADERS,
    );
  });

  it("importTransactions POSTs FormData with files field", async () => {
    mockPost.mockResolvedValue({ files: [], rows_inserted: 0, rows_duplicate: 0 });
    const file = new File(["xls"], "tx.xls", { type: "application/vnd.ms-excel" });

    await gsmApi.importTransactions([file]);

    expect(mockPost).toHaveBeenCalledTimes(1);
    const [url, body] = mockPost.mock.calls[0] as [string, FormData];
    expect(url).toBe("/api/v1/gsm/transactions/import");
    expect(body).toBeInstanceOf(FormData);
    expect(body.getAll("files")).toHaveLength(1);
  });

  it("listWaybills GETs with vehicle and period query", async () => {
    mockGet.mockResolvedValue([]);
    await gsmApi.listWaybills({ vehicleId: 4, periodFrom: "2025-04-01", periodTo: "2025-04-30" });
    expect(mockGet).toHaveBeenCalledWith(
      "/api/v1/gsm/waybills?vehicle_id=4&from=2025-04-01&to=2025-04-30",
    );
  });

  it("generateWaybills POSTs JSON to /waybills/generate", async () => {
    const payload = {
      vehicle_id: 4,
      period_from: "2025-04-01",
      period_to: "2025-04-30",
      force: true,
      fuel_start: 20,
      odometer_start: 10000,
    };
    const generated = {
      waybills: [],
      warnings: [],
      days_created: 0,
      problematic_days: [],
      manual_days: 0,
    };
    mockPost.mockResolvedValue(generated);
    await expect(gsmApi.generateWaybills(payload)).resolves.toEqual(generated);
    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/gsm/waybills/generate",
      JSON.stringify(payload),
      JSON_HEADERS,
    );
  });

  it("listRoutes GETs with vehicle_id", async () => {
    mockGet.mockResolvedValue([]);
    await gsmApi.listRoutes(4);
    expect(mockGet).toHaveBeenCalledWith("/api/v1/gsm/routes?vehicle_id=4");
  });

  it("patchWaybill PATCHes waybill id", async () => {
    mockPatch.mockResolvedValue({ id: 9 });
    await gsmApi.patchWaybill(9, { driver_id: 2, km: 120 });
    expect(mockPatch).toHaveBeenCalledWith(
      "/api/v1/gsm/waybills/9",
      JSON.stringify({ driver_id: 2, km: 120 }),
      JSON_HEADERS,
    );
  });

  it("createWaybill POSTs to /waybills", async () => {
    const payload = {
      vehicle_id: 1,
      date: "2025-04-03",
      driver_id: 7,
      route: [{ from: "A", to: "B", km: 100 }],
      fuel_issued: 0,
    };
    mockPost.mockResolvedValue({ id: 1, ...payload });
    await gsmApi.createWaybill(payload);
    expect(mockPost).toHaveBeenCalledWith(
      "/api/v1/gsm/waybills",
      JSON.stringify(payload),
      JSON_HEADERS,
    );
  });

  it("exportWaybills POSTs JSON via download to /waybills/export", async () => {
    const payload = {
      vehicle_ids: [4],
      from: "2025-04-01",
      to: "2025-04-30",
    };
    const downloaded = {
      blob: new Blob(["zip"], { type: "application/zip" }),
      filename: "gsm_waybills_2025-04-01_2025-04-30.zip",
      contentType: "application/zip",
    };
    mockDownload.mockResolvedValue(downloaded);

    await expect(gsmApi.exportWaybills(payload)).resolves.toEqual(downloaded);
    expect(mockDownload).toHaveBeenCalledWith(
      "/api/v1/gsm/waybills/export",
      "gsm_waybills_2025-04-01_2025-04-30.zip",
      {
        method: "POST",
        body: JSON.stringify(payload),
        headers: JSON_HEADERS,
      },
    );
  });
});
