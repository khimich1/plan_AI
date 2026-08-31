import type { PageSource } from "@/features/commercial-offer/lib/multiPageSource";

/** Shared multi-page source props for *InputStep + SourceInputCard. */
export type MultiPageSourceStepProps = {
  pages: PageSource[];
  activePageId: string | null;
  softCapMessage?: string | null;
  pageProgressLabel?: string | null;
  /** First-source multi; append card stays single-file. */
  singleFileOnly?: boolean;
  /** True after multi-page OCR start — drives wait banner + lightbox lock. */
  recognitionStarted?: boolean;
  onAddFiles: (files: File[]) => void;
  onRemovePage: (id: string) => void;
  onSelectPage: (id: string) => void;
  onPrevPage?: () => void;
  onNextPage?: () => void;
  /** When multi-page review: only ready page can be confirmed. */
  canConfirmActivePage?: boolean;
};
