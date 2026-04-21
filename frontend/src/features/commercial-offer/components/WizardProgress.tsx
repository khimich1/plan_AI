import type { WizardStepId } from "@/features/commercial-offer/types/commercialOffer";
import { Card } from "@/shared/ui/Card";

const steps: Array<{ id: WizardStepId; title: string }> = [
  { id: "plates", title: "1. Ввод плит" },
  { id: "wide-plates", title: "2. Проверка проблемных плит" },
  { id: "manager", title: "3. Менеджер" },
  { id: "client", title: "4. Клиент и условия" },
  { id: "result", title: "5. Расчёт и результат" },
];

type WizardProgressProps = {
  currentStep: WizardStepId;
};

export const WizardProgress = ({ currentStep }: WizardProgressProps) => (
  <Card title="Шаги мастера" subtitle="Состояние сохраняется в браузере до завершения сценария.">
    <div style={{ display: "grid", gap: "0.75rem" }}>
      {steps.map((step, index) => {
        const isActive = step.id === currentStep;
        const isPassed = steps.findIndex((item) => item.id === currentStep) > index;
        return (
          <div
            key={step.id}
            style={{
              borderRadius: 14,
              padding: "0.8rem 0.9rem",
              border: "1px solid #e4e7ec",
              background: isActive ? "#eef2ff" : "#ffffff",
              color: isPassed ? "#067647" : isActive ? "#2b5cff" : "#344054",
              fontWeight: isActive ? 700 : 500,
            }}
          >
            {step.title}
          </div>
        );
      })}
    </div>
  </Card>
);
