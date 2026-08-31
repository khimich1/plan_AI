export const MAX_PAGES = 12;

export type PageStatus = "pending" | "running" | "ready" | "error" | "confirmed";

export type PageSource = {
  id: string;
  file: File;
  name: string;
  previewUrl: string;
  status: PageStatus;
  errorMessage?: string;
  /** blob/remote URL of large preview after OCR, if different from gallery preview */
  recognizedImageUrl?: string | null;
  /** Second OCR pass failed for this page (empty verify list or API exception). */
  ocrVerifyFailed?: boolean;
  batchReviewText: string;
};

export type PageIdFactory = {
  createId: () => string;
  createPreviewUrl: (file: File) => string;
};

export type AddFilesResult =
  | {
      ok: true;
      pages: PageSource[];
      rejectedCount: number;
      rejectReason?: string;
    }
  | {
      ok: false;
      pages: PageSource[];
      rejectedCount: number;
      rejectReason: string;
    };

const softCapRejectReason = () =>
  `Можно добавить не больше ${MAX_PAGES} страниц.`;

export const createPageSource = (file: File, factories: PageIdFactory): PageSource => ({
  id: factories.createId(),
  file,
  name: file.name,
  previewUrl: factories.createPreviewUrl(file),
  status: "pending",
  batchReviewText: "",
});

export const canRemovePage = (page: { status: PageStatus }): boolean =>
  page.status === "pending" || page.status === "error" || page.status === "ready";

export const addFilesToPages = (
  pages: PageSource[],
  files: File[],
  factories: PageIdFactory,
): AddFilesResult => {
  if (files.length === 0) {
    return { ok: true, pages, rejectedCount: 0 };
  }

  const slotsLeft = Math.max(0, MAX_PAGES - pages.length);
  if (slotsLeft === 0) {
    return {
      ok: false,
      pages,
      rejectedCount: files.length,
      rejectReason: softCapRejectReason(),
    };
  }

  const accepted = files.slice(0, slotsLeft);
  const rejectedCount = files.length - accepted.length;
  const appended = accepted.map((file) => createPageSource(file, factories));
  const nextPages = [...pages, ...appended];

  if (rejectedCount > 0) {
    return {
      ok: true,
      pages: nextPages,
      rejectedCount,
      rejectReason: softCapRejectReason(),
    };
  }

  return { ok: true, pages: nextPages, rejectedCount: 0 };
};

export const removePage = (
  pages: PageSource[],
  id: string,
): { pages: PageSource[]; removed: PageSource | null } => {
  const index = pages.findIndex((page) => page.id === id);
  if (index < 0) {
    return { pages, removed: null };
  }
  const removed = pages[index]!;
  return {
    pages: [...pages.slice(0, index), ...pages.slice(index + 1)],
    removed,
  };
};

export const pickNextReadyAfterConfirm = (
  pages: PageSource[],
  confirmedId: string,
): string | null => {
  const confirmedIndex = pages.findIndex((page) => page.id === confirmedId);
  if (confirmedIndex >= 0) {
    for (let i = confirmedIndex + 1; i < pages.length; i += 1) {
      if (pages[i]?.status === "ready") {
        return pages[i]!.id;
      }
    }
  }
  const firstReady = pages.find((page) => page.status === "ready");
  return firstReady?.id ?? null;
};

export const countRecognizedProgress = (
  pages: PageSource[],
): { recognized: number; total: number } => {
  const recognized = pages.filter(
    (page) => page.status === "ready" || page.status === "confirmed",
  ).length;
  return { recognized, total: pages.length };
};

/** Wait UX: show banner only after start and before the first ready/confirmed page. */
export const isWaitingForFirstOcrReady = (
  hasStarted: boolean,
  pages: Array<{ status: PageStatus }>,
): boolean => {
  if (!hasStarted) {
    return false;
  }
  return !pages.some((page) => page.status === "ready" || page.status === "confirmed");
};

export const OCR_WAIT_MESSAGE = "Идёт распознавание, подождите 1–2 минуты";
