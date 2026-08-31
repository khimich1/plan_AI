import { describe, expect, it } from "vitest";

import {
  MAX_PAGES,
  addFilesToPages,
  canRemovePage,
  countRecognizedProgress,
  createPageSource,
  isWaitingForFirstOcrReady,
  pickNextReadyAfterConfirm,
  removePage,
  type PageSource,
  type PageStatus,
} from "@/features/commercial-offer/lib/multiPageSource";

const makeFile = (name: string) => new File(["x"], name, { type: "image/png" });

const page = (
  overrides: Partial<PageSource> & { id: string; status: PageStatus; name?: string },
): PageSource => ({
  id: overrides.id,
  file: overrides.file ?? makeFile(overrides.name ?? `${overrides.id}.png`),
  name: overrides.name ?? `${overrides.id}.png`,
  previewUrl: overrides.previewUrl ?? `blob:${overrides.id}`,
  status: overrides.status,
  errorMessage: overrides.errorMessage,
  recognizedImageUrl: overrides.recognizedImageUrl,
  batchReviewText: overrides.batchReviewText ?? "",
});

describe("MAX_PAGES", () => {
  it("is soft-cap 12", () => {
    expect(MAX_PAGES).toBe(12);
  });
});

describe("canRemovePage", () => {
  it.each(["pending", "error", "ready"] as const)("allows remove for %s", (status) => {
    expect(canRemovePage(page({ id: "a", status }))).toBe(true);
  });

  it.each(["running", "confirmed"] as const)("forbids remove for %s", (status) => {
    expect(canRemovePage(page({ id: "a", status }))).toBe(false);
  });
});

describe("addFilesToPages", () => {
  it("appends files to the end with pending status", () => {
    const existing = [page({ id: "a", status: "ready", batchReviewText: "keep me" })];
    const result = addFilesToPages(existing, [makeFile("b.png"), makeFile("c.png")], {
      createId: () => "new-id",
      createPreviewUrl: (file) => `preview:${file.name}`,
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.pages).toHaveLength(3);
    expect(result.pages[0]).toEqual(existing[0]);
    expect(result.pages[1]?.name).toBe("b.png");
    expect(result.pages[1]?.status).toBe("pending");
    expect(result.pages[1]?.previewUrl).toBe("preview:b.png");
    expect(result.pages[2]?.name).toBe("c.png");
    expect(result.rejectedCount).toBe(0);
  });

  it("accepts the 12th file and rejects the 13th with a reason", () => {
    const existing = Array.from({ length: 11 }, (_, i) =>
      page({ id: `p${i}`, status: "pending", name: `p${i}.png` }),
    );
    const result = addFilesToPages(existing, [makeFile("ok.png"), makeFile("overflow.png")], {
      createId: (() => {
        let n = 0;
        return () => `id-${n++}`;
      })(),
      createPreviewUrl: (file) => `preview:${file.name}`,
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.pages).toHaveLength(12);
    expect(result.pages[11]?.name).toBe("ok.png");
    expect(result.rejectedCount).toBe(1);
    expect(result.rejectReason).toMatch(/12/);
  });

  it("rejects all files when already at soft-cap", () => {
    const existing = Array.from({ length: 12 }, (_, i) =>
      page({ id: `p${i}`, status: "pending", name: `p${i}.png` }),
    );
    const result = addFilesToPages(existing, [makeFile("x.png")], {
      createId: () => "x",
      createPreviewUrl: () => "preview:x",
    });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.pages).toHaveLength(12);
    expect(result.rejectedCount).toBe(1);
    expect(result.rejectReason).toMatch(/12/);
  });
});

describe("removePage", () => {
  it("removes only the target page and leaves neighbors unchanged", () => {
    const pages = [
      page({ id: "a", status: "ready", batchReviewText: "A" }),
      page({ id: "b", status: "error", batchReviewText: "B", errorMessage: "bad" }),
      page({ id: "c", status: "pending", batchReviewText: "C" }),
    ];
    const result = removePage(pages, "b");
    expect(result.removed).toEqual(pages[1]);
    expect(result.pages).toHaveLength(2);
    expect(result.pages[0]).toEqual(pages[0]);
    expect(result.pages[1]).toEqual(pages[2]);
  });

  it("returns null removed when id is missing", () => {
    const pages = [page({ id: "a", status: "pending" })];
    const result = removePage(pages, "missing");
    expect(result.removed).toBeNull();
    expect(result.pages).toEqual(pages);
  });
});

describe("pickNextReadyAfterConfirm", () => {
  it("picks the nearest following ready page after confirmed id", () => {
    const pages = [
      page({ id: "a", status: "confirmed" }),
      page({ id: "b", status: "running" }),
      page({ id: "c", status: "ready" }),
      page({ id: "d", status: "ready" }),
    ];
    expect(pickNextReadyAfterConfirm(pages, "a")).toBe("c");
  });

  it("wraps to the first ready when nothing follows", () => {
    const pages = [
      page({ id: "a", status: "ready" }),
      page({ id: "b", status: "confirmed" }),
      page({ id: "c", status: "pending" }),
    ];
    expect(pickNextReadyAfterConfirm(pages, "b")).toBe("a");
  });

  it("returns null when no ready pages remain", () => {
    const pages = [
      page({ id: "a", status: "confirmed" }),
      page({ id: "b", status: "running" }),
    ];
    expect(pickNextReadyAfterConfirm(pages, "a")).toBeNull();
  });
});

describe("countRecognizedProgress", () => {
  it("counts ready+confirmed as recognized over total pages", () => {
    const pages = [
      page({ id: "a", status: "confirmed" }),
      page({ id: "b", status: "ready" }),
      page({ id: "c", status: "running" }),
      page({ id: "d", status: "error" }),
    ];
    expect(countRecognizedProgress(pages)).toEqual({ recognized: 2, total: 4 });
  });
});

describe("createPageSource", () => {
  it("builds a pending page from a file", () => {
    const file = makeFile("shot.png");
    const created = createPageSource(file, {
      createId: () => "id-1",
      createPreviewUrl: () => "blob:1",
    });
    expect(created).toMatchObject({
      id: "id-1",
      file,
      name: "shot.png",
      previewUrl: "blob:1",
      status: "pending",
      batchReviewText: "",
    });
  });
});

describe("isWaitingForFirstOcrReady", () => {
  it("is false before start even with pending pages", () => {
    expect(isWaitingForFirstOcrReady(false, [page({ id: "a", status: "pending" })])).toBe(false);
  });

  it("is true when started and all pages are pending or running", () => {
    expect(
      isWaitingForFirstOcrReady(true, [
        page({ id: "a", status: "running" }),
        page({ id: "b", status: "pending" }),
      ]),
    ).toBe(true);
  });

  it("is false once any page is ready", () => {
    expect(
      isWaitingForFirstOcrReady(true, [
        page({ id: "a", status: "ready" }),
        page({ id: "b", status: "pending" }),
      ]),
    ).toBe(false);
  });

  it("is false once any page is confirmed", () => {
    expect(
      isWaitingForFirstOcrReady(true, [
        page({ id: "a", status: "confirmed" }),
        page({ id: "b", status: "running" }),
      ]),
    ).toBe(false);
  });
});
