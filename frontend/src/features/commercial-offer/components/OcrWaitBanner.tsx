import { OCR_WAIT_MESSAGE } from "@/features/commercial-offer/lib/multiPageSource";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";

export const OcrWaitBanner = () => (
  <Alert tone="info">
    <div
      role="status"
      aria-live="polite"
      style={{ display: "flex", alignItems: "center", gap: "0.75rem" }}
    >
      <Spinner />
      <span>{OCR_WAIT_MESSAGE}</span>
    </div>
  </Alert>
);
