import type { Manager } from "@/features/commercial-offer/types/commercialOffer";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { StepLayout } from "@/shared/ui/StepLayout";

type ManagerStepProps = {
  managers: Manager[];
  selectedManagerId: number | null;
  errorMessage: string | null;
  isPending: boolean;
  onSelect: (managerId: number) => void;
  onBack: () => void;
  onNext: () => void;
};

export const ManagerStep = ({
  managers,
  selectedManagerId,
  errorMessage,
  isPending,
  onSelect,
  onBack,
  onNext,
}: ManagerStepProps) => {
  const selectedManager = managers.find((item) => item.id === selectedManagerId) ?? null;

  return (
    <StepLayout
      title="Шаг 3. Выбор менеджера"
      description="Список менеджеров загружается из backend API. Выбранный менеджер будет использован в итоговых документах КП."
      footer={
        <div style={{ display: "flex", justifyContent: "space-between", gap: "0.75rem" }}>
          <Button type="button" variant="ghost" onClick={onBack}>
            Назад
          </Button>
          <Button type="button" onClick={onNext} disabled={isPending || !selectedManagerId}>
            {isPending ? "Сохраняем..." : "Далее"}
          </Button>
        </div>
      }
    >
      {errorMessage && <Alert tone="error">{errorMessage}</Alert>}
      <Card title="Менеджеры">
        <div style={{ display: "grid", gap: "0.75rem" }}>
          {managers.map((manager) => (
            <label
              key={manager.id}
              style={{
                display: "grid",
                gap: "0.25rem",
                border: selectedManagerId === manager.id ? "2px solid #2b5cff" : "1px solid #e4e7ec",
                borderRadius: 14,
                padding: "0.9rem",
                background: selectedManagerId === manager.id ? "#eef2ff" : "#ffffff",
              }}
            >
              <div style={{ display: "flex", gap: "0.75rem", alignItems: "center" }}>
                <input type="radio" checked={selectedManagerId === manager.id} onChange={() => onSelect(manager.id)} />
                <strong>{manager.fio}</strong>
              </div>
              <span style={{ color: "#475467" }}>{manager.contact_number || "Телефон не указан"}</span>
              <span style={{ color: "#475467" }}>{manager.email || "Email не указан"}</span>
            </label>
          ))}
        </div>
      </Card>

      {selectedManager && (
        <Card title="Выбранный менеджер">
          <div style={{ display: "grid", gap: "0.35rem" }}>
            <strong>{selectedManager.fio}</strong>
            <span>Телефон: {selectedManager.contact_number || "не указан"}</span>
            <span>Email: {selectedManager.email || "не указан"}</span>
          </div>
        </Card>
      )}
    </StepLayout>
  );
};
