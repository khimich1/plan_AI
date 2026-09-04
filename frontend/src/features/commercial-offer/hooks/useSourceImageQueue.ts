import { useCallback, useEffect, useRef, useState } from "react";

import {
  buildQueueFromPages,
  buildQueueFromPreview,
  clearQueueItems,
  type PageLikeForQueue,
  type PreviewLike,
  type SourceImageQueueItem,
} from "@/features/commercial-offer/lib/sourceImageQueue";

type UrlFactories = {
  createObjectURL?: (file: File) => string;
  revokeObjectURL?: (url: string) => void;
};

export type UseSourceImageQueueResult = {
  items: SourceImageQueueItem[];
  length: number;
  setFromPages: (pages: PageLikeForQueue[]) => void;
  setFromSinglePreview: (source: File | PreviewLike | null | undefined) => void;
  clear: () => void;
};

const defaultCreateObjectURL = (file: File) => URL.createObjectURL(file);
const defaultRevokeObjectURL = (url: string) => URL.revokeObjectURL(url);

export const useSourceImageQueue = (
  factories: UrlFactories = {},
): UseSourceImageQueueResult => {
  const createObjectURLRef = useRef(factories.createObjectURL ?? defaultCreateObjectURL);
  const revokeObjectURLRef = useRef(factories.revokeObjectURL ?? defaultRevokeObjectURL);
  createObjectURLRef.current = factories.createObjectURL ?? defaultCreateObjectURL;
  revokeObjectURLRef.current = factories.revokeObjectURL ?? defaultRevokeObjectURL;

  const [items, setItems] = useState<SourceImageQueueItem[]>([]);
  const itemsRef = useRef<SourceImageQueueItem[]>([]);

  const replaceItems = useCallback((next: SourceImageQueueItem[]) => {
    const previous = itemsRef.current;
    itemsRef.current = next;
    setItems(next);
    clearQueueItems(previous, revokeObjectURLRef.current);
  }, []);

  const setFromPages = useCallback(
    (pages: PageLikeForQueue[]) => {
      replaceItems(buildQueueFromPages(pages, (file) => createObjectURLRef.current(file)));
    },
    [replaceItems],
  );

  const setFromSinglePreview = useCallback(
    (source: File | PreviewLike | null | undefined) => {
      replaceItems(buildQueueFromPreview(source, (file) => createObjectURLRef.current(file)));
    },
    [replaceItems],
  );

  const clear = useCallback(() => {
    replaceItems([]);
  }, [replaceItems]);

  // Revoke only on unmount. Putting URL.revokeObjectURL.bind(URL) in effect
  // deps revokes live queue urls on every parent re-render (broken <img>).
  useEffect(
    () => () => {
      clearQueueItems(itemsRef.current, revokeObjectURLRef.current);
    },
    [],
  );

  return {
    items,
    length: items.length,
    setFromPages,
    setFromSinglePreview,
    clear,
  };
};
