import type { WizardStepId } from "@/features/commercial-offer/types/commercialOffer";
import { Card } from "@/shared/ui/Card";

const steps: Array<{ id: WizardStepId; title: string }> = [
  { id: "plates", title: "1. Плиты" },
  { id: "client", title: "2. Клиент" },
  { id: "result", title: "3. Результат" },
];

type WizardProgressProps = {
  currentStep: WizardStepId;
  onStepClick: (step: WizardStepId) => void;
  canNavigateToStep: (step: WizardStepId) => boolean;
};

export const WizardProgress = ({ currentStep, onStepClick, canNavigateToStep }: WizardProgressProps) => (
  <Card title="Шаги мастера" subtitle="Состояние сохраняется в браузере до завершения сценария.">
    <div style={{ display: "grid", gap: "0.75rem" }}>
      {steps.map((step, index) => {
        const isActive = step.id === currentStep;
        const isPassed = steps.findIndex((item) => item.id === currentStep) > index;
        const isEnabled = canNavigateToStep(step.id);
        return (
          <button
            key={step.id}
            type="button"
            onClick={() => onStepClick(step.id)}
            disabled={!isEnabled}
            style={{
              borderRadius: 14,
              padding: "0.8rem 0.9rem",
              border: "1px solid #e4e7ec",
              background: isActive ? "#eef2ff" : "#ffffff",
              color: isPassed ? "#067647" : isActive ? "#2b5cff" : "#344054",
              fontWeight: isActive ? 700 : 500,
              textAlign: "left",
              cursor: isEnabled ? "pointer" : "not-allowed",
              opacity: isEnabled ? 1 : 0.55,
            }}
          >
            {step.title}
          </button>
        );
      })}
    </div>
  </Card>
);
