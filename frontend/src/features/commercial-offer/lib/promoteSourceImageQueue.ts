import type {
  PageLikeForQueue,
  PreviewLike,
} from "@/features/commercial-offer/lib/sourceImageQueue";

export type SourceImageQueueSink = {
  setFromPages: (pages: PageLikeForQueue[]) => void;
  setFromSinglePreview: (source: File | PreviewLike | null | undefined) => void;
};

export type PromoteSourceImageQueueInput = {
  pages: PageLikeForQueue[];
  singlePreview: File | PreviewLike | null | undefined;
};

export type PromoteSourceImageQueueResult =
  | { promoted: true; kind: "pages" | "single" }
  | { promoted: false; kind: "none" };

/**
 * Snapshot source images into the sticky queue BEFORE OCR reset / preview clear.
 * Multi-page wins over single preview when both are present.
 */
export const applyPromoteSourceImageQueue = (
  sink: SourceImageQueueSink,
  input: PromoteSourceImageQueueInput,
): PromoteSourceImageQueueResult => {
  if (input.pages.length > 0) {
    sink.setFromPages(input.pages);
    return { promoted: true, kind: "pages" };
  }
  if (input.singlePreview) {
    sink.setFromSinglePreview(input.singlePreview);
    return { promoted: true, kind: "single" };
  }
  return { promoted: false, kind: "none" };
};
