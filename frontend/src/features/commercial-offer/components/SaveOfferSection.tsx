import type { CommercialDraftDetails, CommercialSaveResult, SaveMode } from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";

type SaveOfferSectionProps = {
  draft: CommercialDraftDetails;
  lastSaveResult: CommercialSaveResult | null;
  isPending: boolean;
  onSave: (payload: { mode: SaveMode; executionTermsInput: string }) => Promise<void>;
};

export const SaveOfferSection = ({
  draft,
  lastSaveResult,
  isPending,
  onSave,
}: SaveOfferSectionProps) => {
  const resumeKpId = draft.metadata.resume_kp_id ?? null;
  const isResumeEdit = resumeKpId != null && Number(resumeKpId) > 0;
  const savedThisSession = Boolean(lastSaveResult);
  const createAlreadySaved = !isResumeEdit && Boolean(lastSaveResult ?? draft.saved_offer);
  const isSubmitDisabled = isPending || createAlreadySaved || (isResumeEdit && savedThisSession);

  const primaryLabel = (() => {
    if (isPending) {
      return "Сохранение...";
    }
    if (isResumeEdit && !savedThisSession) {
      return "Сохранить изменения";
    }
    if (createAlreadySaved || savedThisSession) {
      return "Сохранено";
    }
    return "В архив";
  })();

  const handlePrimarySave = async () => {
    if (isSubmitDisabled) {
      return;
    }
    await onSave({ mode: "archive", executionTermsInput: "" });
  };

  return (
    <Card
      title="Сохранение результата"
      subtitle={
        isResumeEdit
          ? "Сохраните правки в то же КП — оно останется в архиве."
          : "По умолчанию КП сохраняется в архив — для завершённых предложений."
      }
    >
      <div style={{ display: "grid", gap: "1rem" }}>
        <Button
          type="button"
          variant={isSubmitDisabled ? "secondary" : "primary"}
          disabled={isSubmitDisabled}
          onClick={() => void handlePrimarySave()}
        >
          {primaryLabel}
        </Button>
      </div>

      {(lastSaveResult ?? draft.saved_offer) && (
        <div style={{ marginTop: "1rem" }}>
          <Alert tone="success">
            {lastSaveResult?.result_card.status ?? draft.saved_offer?.status}:{" "}
            {lastSaveResult?.result_card.offer_number ?? draft.offer_identity.offer_number}
          </Alert>
        </div>
      )}
    </Card>
  );
};
