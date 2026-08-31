import { afterEach, describe, expect, it, vi } from "vitest";
import { httpClient } from "@/shared/api/httpClient";
import { queryClient } from "@/shared/lib/queryClient";

describe("httpClient CSRF bootstrap", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    document.cookie = "csrf_token=; Max-Age=0; Path=/";
  });

  it("fetches /api/v1/health and sends X-CSRF-Token when cookie was missing", async () => {
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      if (url.includes("/api/v1/health") && (init?.method ?? "GET") === "GET") {
        document.cookie = "csrf_token=bootstrapped-token; Path=/";
        return new Response(JSON.stringify({ status: "ok" }), {
          status: 200,
          headers: { "Content-Type": "application/json" },
        });
      }
      if (url.includes("/api/v1/auth/login")) {
        expect(init?.headers).toBeTruthy();
        const headers = new Headers(init?.headers);
        expect(headers.get("X-CSRF-Token")).toBe("bootstrapped-token");
        return new Response(JSON.stringify({ user: { id: 1, username: "admin", role: "admin" } }), {
          status: 200,
          headers: { "Content-Type": "application/json" },
        });
      }
      throw new Error(`Unexpected fetch: ${url}`);
    });
    vi.stubGlobal("fetch", fetchMock);

    await httpClient.post(
      "/api/v1/auth/login",
      JSON.stringify({ username: "admin", password: "x" }),
      { "Content-Type": "application/json" },
    );

    expect(fetchMock).toHaveBeenCalled();
    const urls = fetchMock.mock.calls.map((c) => String(c[0]));
    expect(urls.some((u) => u.includes("/api/v1/health"))).toBe(true);
    expect(urls.some((u) => u.includes("/api/v1/auth/login"))).toBe(true);
  });
});

describe("httpClient auth/me 401", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    queryClient.clear();
  });

  it("does not invalidate auth/me on 401 (avoids refetch loop)", async () => {
    const invalidateSpy = vi.spyOn(queryClient, "invalidateQueries");
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response(JSON.stringify({ detail: "Not authenticated" }), {
            status: 401,
            headers: { "Content-Type": "application/json" },
          }),
      ),
    );

    await expect(httpClient.get("/api/v1/auth/me")).rejects.toMatchObject({ status: 401 });
    expect(invalidateSpy).not.toHaveBeenCalled();
  });
});

describe("httpClient.download POST", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    document.cookie = "csrf_token=; Max-Age=0; Path=/";
  });

  it("sends CSRF and returns a blob for POST export", async () => {
    document.cookie = "csrf_token=export-csrf; Path=/";
    const zipBytes = new Uint8Array([0x50, 0x4b]);
    vi.stubGlobal(
      "fetch",
      vi.fn(async (_input: RequestInfo | URL, init?: RequestInit) => {
        expect(init?.method).toBe("POST");
        const headers = new Headers(init?.headers);
        expect(headers.get("X-CSRF-Token")).toBe("export-csrf");
        expect(headers.get("Content-Type")).toBe("application/json");
        return new Response(zipBytes, {
          status: 200,
          headers: {
            "Content-Type": "application/zip",
            "Content-Disposition": 'attachment; filename="gsm_waybills.zip"',
          },
        });
      }),
    );

    const result = await httpClient.download("/api/v1/gsm/waybills/export", "fallback.zip", {
      method: "POST",
      body: JSON.stringify({ vehicle_ids: [1], from: "2025-04-01", to: "2025-04-30" }),
      headers: { "Content-Type": "application/json" },
    });

    expect(result.filename).toBe("gsm_waybills.zip");
    expect(result.contentType).toContain("application/zip");
    expect(new Uint8Array(await result.blob.arrayBuffer())).toEqual(zipBytes);
  });
});
