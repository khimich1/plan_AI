import { useEffect, useMemo, useState } from "react";
import { useSearchParams } from "react-router-dom";
import { CommercialOfferWizard } from "@/features/commercial-offer/components/CommercialOfferWizard";
import { useWizardDraftStore } from "@/features/commercial-offer/store/wizardDraftStore";
import { Alert } from "@/shared/ui/Alert";

const LEGACY_DRAFT_NOTICE =
  "Старый HTML-интерфейс черновиков устарел. Продолжайте работу в новом мастере КП.";

export const CommercialOfferCreatePage = () => {
  const [searchParams, setSearchParams] = useSearchParams();
  const { state, dispatch } = useWizardDraftStore();
  const [dismissedBanner, setDismissedBanner] = useState(false);

  const draftIdFromUrl = searchParams.get("draft")?.trim() ?? "";
  const legacyRedirect = searchParams.get("legacy") === "1";
  const noticeFromUrl = searchParams.get("notice")?.trim() ?? "";
  const errorFromUrl = searchParams.get("error")?.trim() ?? "";

  useEffect(() => {
    if (!draftIdFromUrl || state.draftId === draftIdFromUrl) {
      return;
    }
    dispatch({ type: "set-draft-id", draftId: draftIdFromUrl });
  }, [dispatch, draftIdFromUrl, state.draftId]);

  const infoBannerMessage = useMemo(() => {
    if (dismissedBanner) {
      return null;
    }
    if (noticeFromUrl) {
      return noticeFromUrl;
    }
    if (legacyRedirect && draftIdFromUrl) {
      return LEGACY_DRAFT_NOTICE;
    }
    return null;
  }, [dismissedBanner, draftIdFromUrl, legacyRedirect, noticeFromUrl]);

  const dismissInfoBanner = () => {
    setDismissedBanner(true);
    const next = new URLSearchParams(searchParams);
    next.delete("legacy");
    next.delete("notice");
    setSearchParams(next, { replace: true });
  };

  const dismissErrorBanner = () => {
    const next = new URLSearchParams(searchParams);
    next.delete("error");
    setSearchParams(next, { replace: true });
  };

  return (
    <main style={{ maxWidth: 1280, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
      {errorFromUrl ? (
        <div style={{ marginBottom: "1rem" }}>
          <Alert tone="error">
            <div style={{ display: "flex", justifyContent: "space-between", gap: "1rem", alignItems: "flex-start" }}>
              <span>{errorFromUrl}</span>
              <button
                type="button"
                onClick={dismissErrorBanner}
                style={{
                  border: "none",
                  background: "transparent",
                  cursor: "pointer",
                  textDecoration: "underline",
                  whiteSpace: "nowrap",
                }}
              >
                Закрыть
              </button>
            </div>
          </Alert>
        </div>
      ) : null}
      {infoBannerMessage ? (
        <div style={{ marginBottom: "1rem" }}>
          <Alert tone="info">
            <div style={{ display: "flex", justifyContent: "space-between", gap: "1rem", alignItems: "flex-start" }}>
              <span>{infoBannerMessage}</span>
              <button
                type="button"
                onClick={dismissInfoBanner}
                style={{
                  border: "none",
                  background: "transparent",
                  cursor: "pointer",
                  textDecoration: "underline",
                  whiteSpace: "nowrap",
                }}
              >
                Закрыть
              </button>
            </div>
          </Alert>
        </div>
      ) : null}
      <CommercialOfferWizard />
    </main>
  );
};
