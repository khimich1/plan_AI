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
