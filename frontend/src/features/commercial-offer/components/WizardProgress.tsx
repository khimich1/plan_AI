import type { ProductType, WizardStepId } from "@/features/commercial-offer/types/commercialOffer";
import { getProductTypeConfig } from "@/features/commercial-offer/lib/productTypeConfig";
import { getWizardStepOrder } from "@/features/commercial-offer/lib/wizardStepOrder";
import { Card } from "@/shared/ui/Card";

type WizardProgressProps = {
  productType: ProductType;
  currentStep: WizardStepId;
  onStepClick: (step: WizardStepId) => void;
  canNavigateToStep: (step: WizardStepId) => boolean;
  skipClient?: boolean;
};

const stepTitle = (id: WizardStepId): string => {
  if (id === "client") {
    return "2. Клиент";
  }
  if (id === "result") {
    return "3. Результат";
  }
  return `1. ${getProductTypeConfig(id).labels.nounPlural}`;
};

export const WizardProgress = ({
  productType,
  currentStep,
  onStepClick,
  canNavigateToStep,
  skipClient = false,
}: WizardProgressProps) => {
  const steps = getWizardStepOrder(productType, { skipClient }).map((id) => ({
    id,
    title: stepTitle(id),
  }));

  return (
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
};
