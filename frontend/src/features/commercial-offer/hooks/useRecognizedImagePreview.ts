import { useCallback, useEffect, useState } from "react";

export type RecognizedImagePreview = {
  url: string;
  name: string;
};

export const useRecognizedImagePreview = () => {
  const [preview, setPreview] = useState<RecognizedImagePreview | null>(null);

  const setPreviewFromFile = useCallback((file: File) => {
    setPreview((previous) => {
      if (previous) URL.revokeObjectURL(previous.url);
      return { url: URL.createObjectURL(file), name: file.name };
    });
  }, []);

  const clearPreview = useCallback(() => {
    setPreview((previous) => {
      if (previous) URL.revokeObjectURL(previous.url);
      return null;
    });
  }, []);

  useEffect(
    () => () => {
      if (preview) URL.revokeObjectURL(preview.url);
    },
    [preview],
  );

  return { preview, setPreviewFromFile, clearPreview };
};
