export type SourceImageQueueItem = {
  id: string;
  url: string;
  name: string;
};

export type PageLikeForQueue = {
  id: string;
  file: File;
  name: string;
};

export type PreviewLike = {
  url: string;
  name: string;
  /** When set, queue builds an independent object URL from this file. */
  file?: File;
};

type CreateUrl = (file: File) => string;
type RevokeUrl = (url: string) => void;

const SINGLE_PREVIEW_ID = "preview";
const SINGLE_FILE_ID = "single";

export const buildQueueFromPages = (
  pages: PageLikeForQueue[],
  createUrl: CreateUrl,
): SourceImageQueueItem[] =>
  pages.map((page) => ({
    id: page.id,
    url: createUrl(page.file),
    name: page.name,
  }));

export const buildQueueFromPreview = (
  source: File | PreviewLike | null | undefined,
  createUrl?: CreateUrl,
): SourceImageQueueItem[] => {
  if (!source) {
    return [];
  }
  if (source instanceof File) {
    if (!createUrl) {
      throw new Error("createUrl is required when building queue from File");
    }
    return [
      {
        id: SINGLE_FILE_ID,
        url: createUrl(source),
        name: source.name,
      },
    ];
  }
  // Prefer File so the queue owns an independent blob URL (safe if preview URL is revoked).
  if (source.file) {
    if (!createUrl) {
      throw new Error("createUrl is required when building queue from PreviewLike.file");
    }
    return [
      {
        id: SINGLE_PREVIEW_ID,
        url: createUrl(source.file),
        name: source.name || source.file.name,
      },
    ];
  }
  return [
    {
      id: SINGLE_PREVIEW_ID,
      url: source.url,
      name: source.name,
    },
  ];
};

export const clearQueueItems = (
  items: SourceImageQueueItem[],
  revoke: RevokeUrl,
): SourceImageQueueItem[] => {
  for (const item of items) {
    revoke(item.url);
  }
  return [];
};

export const clampQueueIndex = (index: number, length: number): number => {
  if (length <= 0) {
    return 0;
  }
  if (index < 0) {
    return 0;
  }
  if (index >= length) {
    return length - 1;
  }
  return index;
};

export const nextIndex = (index: number, length: number): number =>
  clampQueueIndex(index + 1, length);

export const prevIndex = (index: number, length: number): number =>
  clampQueueIndex(index - 1, length);
