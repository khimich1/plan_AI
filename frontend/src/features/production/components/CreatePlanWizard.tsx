import { Card } from "@/shared/ui/Card";
import { Step1PlanStartDate } from "@/features/production/components/create-plan-wizard/Step1PlanStartDate";
import { Step2TracksConfig } from "@/features/production/components/create-plan-wizard/Step2TracksConfig";
import { Step3KpPlateSelection } from "@/features/production/components/create-plan-wizard/Step3KpPlateSelection";
import { WizardStepIndicator } from "@/features/production/components/create-plan-wizard/WizardStepIndicator";
import { useCreatePlanWizardState } from "@/features/production/hooks/useCreatePlanWizardState";
import type { FillTargetItem } from "@/features/production/types/production";

type Props = {
  onCreated?: () => void;
  /** Если задано — мастер открывается сразу на шаге 3 в режиме «дозаполнение». */
  fillRequest?: FillTargetItem[] | null;
  /** Сигнализирует родителю, что fillRequest подхвачен и можно его очистить. */
  onFillRequestConsumed?: () => void;
  /** Возврат к календарю при отмене дозаполнения. */
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
      {!wizard.isFillMode && <WizardStepIndicator step={wizard.step} />}

      {!wizard.isFillMode && wizard.step === 1 && (
        <Step1PlanStartDate
          startDate={wizard.startDate}
          planName={wizard.planName}
          calendarMonth={wizard.calendarMonth}
          daysInfo={wizard.daysInfo}
          holidays={wizard.holidays}
          extraWorkdays={wizard.extraWorkdays}
          occupiedOnStart={wizard.occupiedOnStart}
          maxPerDay={wizard.maxPerDay}
          freeOnStart={wizard.freeOnStart}
          calendarLoading={wizard.calendarLoading}
          canProceed={wizard.canProceedStep1}
          onStartDateChange={wizard.setStartDate}
          onPlanNameChange={wizard.setPlanName}
          onCalendarMonthChange={wizard.setCalendarMonth}
          onNext={() => wizard.setStep(2)}
        />
      )}

      {!wizard.isFillMode && wizard.step === 2 && (
        <Step2TracksConfig
          tracksCount={wizard.tracksCount}
          canProceed={wizard.canProceedStep2}
          onTracksCountChange={wizard.setTracksCount}
          onBack={() => wizard.setStep(1)}
          onNext={() => wizard.setStep(3)}
        />
      )}

      {wizard.step === 3 && (
        <Step3KpPlateSelection
          isFillMode={wizard.isFillMode}
          planName={wizard.planName}
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
          buildPending={wizard.buildMutation.isPending}
          buildSuccess={wizard.buildMutation.isSuccess}
          buildErrorMessage={wizard.buildErrorMessage}
          canSubmit={wizard.canSubmit}
          onPlanNameChange={wizard.setPlanName}
          onFilterMethodChange={wizard.setFilterMethod}
          onToggleKp={wizard.toggleKp}
          onToggleExpand={wizard.toggleExpand}
          onTogglePlate={wizard.togglePlate}
          onSetPlateQty={wizard.setPlateQty}
          onBack={() => wizard.setStep(2)}
          onCancelFill={wizard.handleCancelFill}
          onSubmit={wizard.handleSubmit}
        />
      )}
    </Card>
  );
};
