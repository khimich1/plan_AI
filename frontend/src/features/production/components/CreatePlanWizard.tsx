import { Card } from "@/shared/ui/Card";
import { Step3KpPlateSelection } from "@/features/production/components/create-plan-wizard/Step3KpPlateSelection";
import { useCreatePlanWizardState } from "@/features/production/hooks/useCreatePlanWizardState";
import type { FillTargetItem } from "@/features/production/types/production";

type Props = {
  onCreated?: () => void;
  /** Корзина с календаря — обязательный вход; без неё ProductionPage редиректит. */
  fillRequest?: FillTargetItem[] | null;
  /** Сигнализирует родителю, что fillRequest подхвачен и можно его очистить. */
  onFillRequestConsumed?: () => void;
  /** Возврат к календарю при отмене. */
  onCancelFill?: () => void;
};

export const CreatePlanWizard = ({
  onCreated,
  fillRequest,
  onFillRequestConsumed,
  onCancelFill,
}: Props) => {
  const wizard = useCreatePlanWizardState({
    onCreated,
    fillRequest,
    onFillRequestConsumed,
    onCancelFill,
  });

  return (
    <Card title={wizard.cardTitle} subtitle={wizard.cardSubtitle}>
      <Step3KpPlateSelection
        isFillMode={wizard.isFillMode}
        filterMethod={wizard.filterMethod}
        candidatesLoading={wizard.candidatesQuery.isLoading}
        candidatesError={wizard.candidatesQuery.isError}
        candidates={wizard.candidatesQuery.data?.items}
        selectionEstimate={wizard.selectionEstimate}
        tracksPerDay={wizard.tracksPerDay}
        tracksPerDaySource={wizard.tracksPerDaySource}
        estimateByKpId={wizard.estimateByKpId}
        selectedPlatesByKp={wizard.selectedPlatesByKp}
        selectedPlateQtyByKp={wizard.selectedPlateQtyByKp}
        expandedKpIds={wizard.expandedKpIds}
        freeQtyByPlateKey={wizard.freeQtyByPlateKey}
        sgpReservationsCount={wizard.sgpReservations.length}
        pendingClose={wizard.pendingClose}
        buildPending={wizard.buildMutation.isPending}
        buildSuccess={wizard.buildMutation.isSuccess}
        buildErrorMessage={wizard.buildErrorMessage}
        canSubmit={wizard.canSubmit}
        urgentPositions={wizard.analyzeResult?.urgent_positions ?? []}
        substrateRecommendations={
          wizard.analyzeResult?.substrate_recommendations ?? []
        }
        capacityDeficit={wizard.analyzeResult?.capacity_deficit ?? null}
        analyzePending={wizard.analyzePending}
        analyzeErrorMessage={wizard.analyzeErrorMessage}
        substrateErrorMessage={wizard.substrateErrorMessage}
        onFilterMethodChange={wizard.setFilterMethod}
        onToggleKp={wizard.toggleKp}
        onToggleExpand={wizard.toggleExpand}
        onTogglePlate={wizard.togglePlate}
        onSetPlateQty={wizard.setPlateQty}
        onProposeCloseFromSgp={(kp, plateId) => {
          const plate = kp.plates.find((p) => p.id === plateId);
          if (plate) wizard.proposeCloseFromSgp(kp, plate);
        }}
        onConfirmCloseFromSgp={wizard.confirmCloseFromSgp}
        onCancelCloseFromSgp={wizard.cancelCloseFromSgp}
        onToggleUrgentPosition={wizard.toggleUrgentPosition}
        onAnalyzeSubstrates={() => wizard.runAnalyzeSubstrates()}
        onToggleSubstrateRecommendation={wizard.toggleSubstrateRecommendation}
        onApplyCapacityOption={wizard.applyCapacityOption}
        onCancelFill={wizard.handleCancelFill}
        onSubmit={wizard.handleSubmit}
      />
    </Card>
  );
};
