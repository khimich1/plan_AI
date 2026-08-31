import { useCallback, useEffect, useRef, useState } from "react";

import {
  addFilesToPages,
  canRemovePage,
  countRecognizedProgress,
  pickNextReadyAfterConfirm,
  removePage,
  type PageSource,
} from "@/features/commercial-offer/lib/multiPageSource";
import type { CommercialDraftDetails, ProductType } from "@/features/commercial-offer/types/commercialOffer";
import { getErrorMessage } from "@/shared/lib/apiError";

export type RecognizePageArgs = {
  image: File;
  productType: ProductType;
  draftId: string | null;
  isFirst: boolean;
};

export type RecognizePageResult = {
  draft: CommercialDraftDetails;
  batchReviewText: string;
};

export type UseMultiPageRecognizeOptions = {
  recognizePage: (args: RecognizePageArgs) => Promise<RecognizePageResult>;
  createId?: () => string;
  createPreviewUrl?: (file: File) => string;
  revokePreviewUrl?: (url: string) => void;
};

let pageIdSeq = 0;

const defaultCreateId = () => {
  pageIdSeq += 1;
  return `page-${pageIdSeq}-${Date.now()}`;
};

const defaultCreatePreviewUrl = (file: File) => URL.createObjectURL(file);
const defaultRevokePreviewUrl = (url: string) => URL.revokeObjectURL(url);

export const useMultiPageRecognize = (options: UseMultiPageRecognizeOptions) => {
  const recognizePageRef = useRef(options.recognizePage);
  recognizePageRef.current = options.recognizePage;

  const createIdRef = useRef(options.createId ?? defaultCreateId);
  createIdRef.current = options.createId ?? defaultCreateId;
  const createPreviewUrlRef = useRef(options.createPreviewUrl ?? defaultCreatePreviewUrl);
  createPreviewUrlRef.current = options.createPreviewUrl ?? defaultCreatePreviewUrl;
  const revokePreviewUrlRef = useRef(options.revokePreviewUrl ?? defaultRevokePreviewUrl);
  revokePreviewUrlRef.current = options.revokePreviewUrl ?? defaultRevokePreviewUrl;

  const [pages, setPages] = useState<PageSource[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);
  const [softCapMessage, setSoftCapMessage] = useState<string | null>(null);
  const [isRunning, setIsRunning] = useState(false);
  const [hasStarted, setHasStarted] = useState(false);
  const [draftId, setDraftId] = useState<string | null>(null);
  const [lastDraft, setLastDraft] = useState<CommercialDraftDetails | null>(null);

  const pagesRef = useRef<PageSource[]>([]);
  const activeIdRef = useRef<string | null>(null);
  const draftIdRef = useRef<string | null>(null);
  const hasStartedRef = useRef(false);
  const productTypeRef = useRef<ProductType | null>(null);
  const createdReplaceRef = useRef(false);
  const runnerLockRef = useRef(false);

  const commitPages = useCallback((next: PageSource[]) => {
    pagesRef.current = next;
    setPages(next);
  }, []);

  const revokePageUrls = useCallback((page: PageSource) => {
    revokePreviewUrlRef.current(page.previewUrl);
    if (page.recognizedImageUrl && page.recognizedImageUrl !== page.previewUrl) {
      revokePreviewUrlRef.current(page.recognizedImageUrl);
    }
  }, []);

  const pumpQueue = useCallback(async () => {
    const productType = productTypeRef.current;
    if (!productType || runnerLockRef.current) {
      return;
    }
    if (!pagesRef.current.some((page) => page.status === "pending")) {
      return;
    }

    runnerLockRef.current = true;
    setIsRunning(true);

    try {
      while (pagesRef.current.some((page) => page.status === "pending")) {
        const pending = pagesRef.current.find((page) => page.status === "pending");
        if (!pending) {
          break;
        }

        const pageId = pending.id;
        const isFirst = !createdReplaceRef.current;

        commitPages(
          pagesRef.current.map((page) =>
            page.id === pageId
              ? { ...page, status: "running" as const, errorMessage: undefined }
              : page,
          ),
        );

        try {
          const result = await recognizePageRef.current({
            image: pending.file,
            productType,
            draftId: draftIdRef.current,
            isFirst,
          });

          createdReplaceRef.current = true;
          setDraftId(result.draft.draft_id);
          draftIdRef.current = result.draft.draft_id;
          setLastDraft(result.draft);

          // R6: reuse gallery previewUrl — do not create a second object URL
          const withReady = pagesRef.current.map((page) =>
            page.id === pageId
              ? {
                  ...page,
                  status: "ready" as const,
                  batchReviewText: result.batchReviewText,
                  recognizedImageUrl: page.previewUrl,
                  errorMessage: undefined,
                  ocrVerifyFailed: Boolean(result.draft.metadata.ocr_verify_failed),
                }
              : page,
          );
          commitPages(withReady);

          const active = withReady.find((page) => page.id === activeIdRef.current);
          const shouldFocus =
            !active ||
            active.status === "pending" ||
            active.status === "running" ||
            active.status === "error";
          if (shouldFocus) {
            setActiveId(pageId);
            activeIdRef.current = pageId;
          }
        } catch (error) {
          const message = getErrorMessage(error);
          commitPages(
            pagesRef.current.map((page) =>
              page.id === pageId ? { ...page, status: "error" as const, errorMessage: message } : page,
            ),
          );
        }
      }
    } finally {
      runnerLockRef.current = false;
      setIsRunning(false);
      if (
        hasStartedRef.current &&
        pagesRef.current.some((page) => page.status === "pending")
      ) {
        void pumpQueue();
      }
    }
  }, [commitPages]);

  const clearSessionFlags = useCallback(() => {
    setSoftCapMessage(null);
    setIsRunning(false);
    setHasStarted(false);
    hasStartedRef.current = false;
    productTypeRef.current = null;
    setDraftId(null);
    draftIdRef.current = null;
    setLastDraft(null);
    createdReplaceRef.current = false;
    runnerLockRef.current = false;
  }, []);

  const addFiles = useCallback(
    (files: File[]): PageSource[] => {
      if (files.length === 0) {
        return pagesRef.current;
      }
      const result = addFilesToPages(pagesRef.current, files, {
        createId: () => createIdRef.current(),
        createPreviewUrl: (file) => createPreviewUrlRef.current(file),
      });
      setSoftCapMessage(result.rejectReason ?? null);
      const previousLength = pagesRef.current.length;
      commitPages(result.pages);
      if (result.pages.length > previousLength && !activeIdRef.current) {
        const firstNew = result.pages[previousLength];
        if (firstNew) {
          setActiveId(firstNew.id);
          activeIdRef.current = firstNew.id;
        }
      }
      if (hasStartedRef.current) {
        void pumpQueue();
      }
      return result.pages;
    },
    [commitPages, pumpQueue],
  );

  const remove = useCallback(
    (id: string): PageSource[] => {
      const target = pagesRef.current.find((page) => page.id === id);
      if (!target || !canRemovePage(target)) {
        return pagesRef.current;
      }
      const { pages: next, removed } = removePage(pagesRef.current, id);
      if (removed) {
        revokePageUrls(removed);
      }
      commitPages(next);
      if (next.length === 0) {
        setActiveId(null);
        activeIdRef.current = null;
        // R1: empty after start must not leave eternal multiPendingReview
        if (hasStartedRef.current) {
          clearSessionFlags();
        }
        return next;
      }
      if (activeIdRef.current === id) {
        const fallback =
          next.find((page) => page.status === "ready") ??
          next.find((page) => page.status === "pending" || page.status === "error") ??
          next[0];
        setActiveId(fallback?.id ?? null);
        activeIdRef.current = fallback?.id ?? null;
      }
      return next;
    },
    [clearSessionFlags, commitPages, revokePageUrls],
  );

  const setActive = useCallback((id: string) => {
    setActiveId(id);
    activeIdRef.current = id;
  }, []);

  const updatePageText = useCallback(
    (id: string, text: string) => {
      commitPages(
        pagesRef.current.map((page) => (page.id === id ? { ...page, batchReviewText: text } : page)),
      );
    },
    [commitPages],
  );

  const start = useCallback(
    async ({
      productType,
      existingDraftId = null,
    }: {
      productType: ProductType;
      existingDraftId?: string | null;
    }) => {
      if (pagesRef.current.length === 0) {
        return;
      }
      productTypeRef.current = productType;
      setHasStarted(true);
      hasStartedRef.current = true;
      if (existingDraftId) {
        setDraftId(existingDraftId);
        draftIdRef.current = existingDraftId;
        createdReplaceRef.current = true;
      } else {
        createdReplaceRef.current = false;
      }
      void pumpQueue();
    },
    [pumpQueue],
  );

  const confirmActive = useCallback((): {
    allConfirmed: boolean;
    nextId: string | null;
    nextBatchReviewText: string;
  } => {
    const currentActiveId = activeIdRef.current;
    if (!currentActiveId) {
      return { allConfirmed: false, nextId: null, nextBatchReviewText: "" };
    }
    const active = pagesRef.current.find((page) => page.id === currentActiveId);
    if (!active || active.status !== "ready") {
      return { allConfirmed: false, nextId: null, nextBatchReviewText: "" };
    }

    const nextPages = pagesRef.current.map((page) =>
      page.id === currentActiveId ? { ...page, status: "confirmed" as const } : page,
    );
    commitPages(nextPages);

    const nextId = pickNextReadyAfterConfirm(nextPages, currentActiveId);
    setActiveId(nextId);
    activeIdRef.current = nextId;
    const nextPage = nextId ? nextPages.find((page) => page.id === nextId) : undefined;
    const confirmedAll = nextPages.length > 0 && nextPages.every((page) => page.status === "confirmed");
    return {
      allConfirmed: confirmedAll,
      nextId,
      nextBatchReviewText: nextPage?.batchReviewText ?? "",
    };
  }, [commitPages]);

  const reset = useCallback(() => {
    pagesRef.current.forEach(revokePageUrls);
    commitPages([]);
    setActiveId(null);
    activeIdRef.current = null;
    clearSessionFlags();
  }, [clearSessionFlags, commitPages, revokePageUrls]);

  useEffect(() => {
    const onBeforeUnload = (event: BeforeUnloadEvent) => {
      if (!hasStartedRef.current) {
        return;
      }
      const busy = pagesRef.current.some(
        (page) => page.status === "pending" || page.status === "running",
      );
      if (!busy) {
        return;
      }
      event.preventDefault();
      event.returnValue = "";
    };
    window.addEventListener("beforeunload", onBeforeUnload);
    return () => window.removeEventListener("beforeunload", onBeforeUnload);
  }, []);

  useEffect(
    () => () => {
      pagesRef.current.forEach(revokePageUrls);
    },
    [revokePageUrls],
  );

  const activePage = pages.find((page) => page.id === activeId) ?? null;
  const progress = countRecognizedProgress(pages);
  const allConfirmed = pages.length > 0 && pages.every((page) => page.status === "confirmed");
  const hasPendingOrRunning = pages.some(
    (page) => page.status === "pending" || page.status === "running",
  );
  // R8: pending pages before start must not look like OCR is running
  const isRecognizing = isRunning || (hasStarted && hasPendingOrRunning);

  return {
    pages,
    activeId,
    activePage,
    softCapMessage,
    isRunning,
    isRecognizing,
    hasStarted,
    draftId,
    lastDraft,
    progress,
    allConfirmed,
    hasPendingOrRunning,
    addFiles,
    remove,
    setActive,
    updatePageText,
    start,
    confirmActive,
    reset,
  };
};
