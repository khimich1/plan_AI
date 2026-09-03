import { useCallback, useEffect, useRef, useState } from "react";

export type RecognizedImagePreview = {
  url: string;
  name: string;
  /** Kept so promote can create an independent object URL (avoid revoke races). */
  file: File;
};

export const useRecognizedImagePreview = () => {
  const [preview, setPreview] = useState<RecognizedImagePreview | null>(null);
  const previewRef = useRef(preview);
  previewRef.current = preview;

  const setPreviewFromFile = useCallback((file: File) => {
    setPreview((previous) => {
      if (previous) URL.revokeObjectURL(previous.url);
      return { url: URL.createObjectURL(file), name: file.name, file };
    });
  }, []);

  const clearPreview = useCallback(() => {
    setPreview((previous) => {
      if (previous) URL.revokeObjectURL(previous.url);
      return null;
    });
  }, []);

  /**
   * Clear preview state and return the File for promote.
   * Revokes the display object URL — caller must create a fresh URL from `file`.
   */
  const takePreview = useCallback((): RecognizedImagePreview | null => {
    const current = previewRef.current;
    if (!current) {
      return null;
    }
    previewRef.current = null;
    setPreview(null);
    URL.revokeObjectURL(current.url);
    return { url: "", name: current.name, file: current.file };
  }, []);

  useEffect(
    () => () => {
      if (previewRef.current) URL.revokeObjectURL(previewRef.current.url);
    },
    [],
  );

  return { preview, setPreviewFromFile, clearPreview, takePreview };
};
