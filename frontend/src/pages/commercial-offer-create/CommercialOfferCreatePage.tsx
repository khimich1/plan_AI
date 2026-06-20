import { useEffect, useMemo, useState } from "react";
import { useSearchParams } from "react-router-dom";
import { CommercialOfferWizard } from "@/features/commercial-offer/components/CommercialOfferWizard";
import { useWizardDraftStore } from "@/features/commercial-offer/store/wizardDraftStore";
import { Alert } from "@/shared/ui/Alert";

const LEGACY_DRAFT_NOTICE =
  "Старый HTML-интерфейс черновиков снят. Продолжите работу в новом мастере КП.";

export const CommercialOfferCreatePage = () => {
  const [searchParams, setSearchParams] = useSearchParams();
  const { state, dispatch } = useWizardDraftStore();
  const [dismissedNotice, setDismissedNotice] = useState(false);

  const draftIdFromUrl = searchParams.get("draft")?.trim() ?? "";
  const legacyRedirect = searchParams.get("legacy") === "1";
  const noticeFromUrl = searchParams.get("notice")?.trim() ?? "";

  useEffect(() => {
    if (!draftIdFromUrl || state.draftId === draftIdFromUrl) {
      return;
    }
    dispatch({ type: "set-draft-id", draftId: draftIdFromUrl });
  }, [dispatch, draftIdFromUrl, state.draftId]);

  const bannerMessage = useMemo(() => {
    if (dismissedNotice) {
      return null;
    }
    if (noticeFromUrl) {
      return noticeFromUrl;
    }
    if (legacyRedirect && draftIdFromUrl) {
      return LEGACY_DRAFT_NOTICE;
    }
    return null;
  }, [dismissedNotice, draftIdFromUrl, legacyRedirect, noticeFromUrl]);

  const dismissBanner = () => {
    setDismissedNotice(true);
    const next = new URLSearchParams(searchParams);
    next.delete("legacy");
    next.delete("notice");
    setSearchParams(next, { replace: true });
  };

  return (
    <main style={{ maxWidth: 1280, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
      {bannerMessage ? (
        <div style={{ marginBottom: "1rem" }}>
          <Alert tone="info">
            <div style={{ display: "flex", justifyContent: "space-between", gap: "1rem", alignItems: "flex-start" }}>
              <span>{bannerMessage}</span>
              <button
                type="button"
                onClick={dismissBanner}
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
