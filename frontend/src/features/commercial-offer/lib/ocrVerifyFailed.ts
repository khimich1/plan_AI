export const OCR_VERIFY_FAILED_REVIEW_MESSAGE =
  "Вторая проверка не подтвердила список — оставили первый вариант. Сверьте с фото.";

export const resolveActivePageOcrVerifyFailed = (
  pages: Array<{ id: string; ocrVerifyFailed?: boolean }>,
  activePageId: string | null,
  draftOcrVerifyFailed?: boolean,
): boolean => {
  if (pages.length === 0) {
    return Boolean(draftOcrVerifyFailed);
  }
  const active = pages.find((page) => page.id === activePageId);
  return Boolean(active?.ocrVerifyFailed);
};

export const resolveActivePageOcrCorrections = <T,>(
  pages: Array<{ id: string; ocrCorrections?: T[] }>,
  activePageId: string | null,
  draftFallback?: T[],
): T[] => {
  if (pages.length === 0) {
    return draftFallback ?? [];
  }
  const active = pages.find((page) => page.id === activePageId);
  if (!active) {
    return [];
  }
  if (active.ocrCorrections !== undefined) {
    return active.ocrCorrections;
  }
  return draftFallback ?? [];
};
