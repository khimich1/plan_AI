import { useEffect, useState } from "react";
import { commercialOfferApi } from "@/features/commercial-offer/api/commercialOfferApi";
import type {
  CommercialParseLine,
  ProductType,
} from "@/features/commercial-offer/types/commercialOffer";

export const SOURCE_TEXT_LINT_DEBOUNCE_MS = 500;

export type SourceLintLine = CommercialParseLine;

type UseSourceTextLintArgs = {
  text: string;
  productType: ProductType;
  enabled: boolean;
};

export type SourceTextLintState = {
  lines: SourceLintLine[];
  isPending: boolean;
  isError: boolean;
};

const EMPTY_STATE: SourceTextLintState = { lines: [], isPending: false, isError: false };

export const useSourceTextLint = ({
  text,
  productType,
  enabled,
}: UseSourceTextLintArgs): SourceTextLintState => {
  const shouldLint = enabled && text.trim().length > 0;
  const [state, setState] = useState<SourceTextLintState>(
    shouldLint ? { ...EMPTY_STATE, isPending: true } : EMPTY_STATE,
  );
  const [trackedKey, setTrackedKey] = useState(shouldLint ? `${productType}\0${text}` : "");
  const currentKey = shouldLint ? `${productType}\0${text}` : "";

  if (currentKey !== trackedKey) {
    setTrackedKey(currentKey);
    setState(shouldLint ? { lines: state.lines, isPending: true, isError: false } : EMPTY_STATE);
  }

  useEffect(() => {
    if (!shouldLint) {
      return;
    }

    const controller = new AbortController();
    let cancelled = false;
    const timer = window.setTimeout(() => {
      void commercialOfferApi
        .parseSource({ text, productType }, { signal: controller.signal })
        .then((response) => {
          if (cancelled) {
            return;
          }
          setState({ lines: response.lines ?? [], isPending: false, isError: false });
        })
        .catch((err: unknown) => {
          if (cancelled || controller.signal.aborted) {
            return;
          }
          if (err instanceof DOMException && err.name === "AbortError") {
            return;
          }
          if (err instanceof Error && err.name === "AbortError") {
            return;
          }
          setState((current) => ({ ...current, isPending: false, isError: true }));
        });
    }, SOURCE_TEXT_LINT_DEBOUNCE_MS);

    return () => {
      cancelled = true;
      controller.abort();
      window.clearTimeout(timer);
    };
  }, [shouldLint, text, productType]);

  return shouldLint ? state : EMPTY_STATE;
};
