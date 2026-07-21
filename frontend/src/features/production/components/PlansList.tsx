import { useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Card } from "@/shared/ui/Card";
import { Modal } from "@/shared/ui/Modal";
import { Spinner } from "@/shared/ui/Spinner";
import { isPlanVersionConflict } from "@/shared/lib/planConflict";
import {
  useActivatePlanMutation,
  useDeletePlanMutation,
  usePlansListQuery,
} from "@/features/production/hooks/useProductionQueries";
import type { PlanMetaSummary } from "@/features/production/types/production";

type Props = {
  onOpenPlanCalendar?: (planId: string) => void;
};

const formatRu = (value: unknown): string => {
  if (!value || typeof value !== "string") return "—";
  const iso = /^\d{4}-\d{2}-\d{2}$/.test(value);
  if (!iso) return value;
  const [y, m, d] = value.split("-");
  return `${d}.${m}.${y}`;
};

export const PlansList = ({ onOpenPlanCalendar }: Props) => {
  const plansQuery = usePlansListQuery();
  const activateMutation = useActivatePlanMutation();
  const deleteMutation = useDeletePlanMutation();
  const [confirmDelete, setConfirmDelete] = useState<PlanMetaSummary | null>(null);

  if (plansQuery.isLoading) {
    return (
      <Card title="Планы">
        <div style={{ display: "flex", gap: "0.5rem" }}>
          <Spinner /> Загрузка…
        </div>
      </Card>
    );
  }

  if (plansQuery.isError) {
    return (
      <Card title="Планы">
        <Alert tone="error">Не удалось загрузить список планов.</Alert>
      </Card>
    );
  }

  const plans = plansQuery.data?.plans ?? [];
  const activeId = plansQuery.data?.active_plan_id ?? null;

  return (
    <Card
      title="Планы"
      subtitle="Список всех сохранённых планов. Звезда обозначает активный план."
    >
      {plans.length === 0 && (
        <Alert tone="info">Нет сохранённых планов. Создайте новый на вкладке «Начать планирование».</Alert>
      )}

      <div style={{ display: "grid", gap: "0.75rem" }}>
        {plans.map((plan) => {
          const isActive = plan.id === activeId;
          return (
            <div
              key={plan.id}
              style={{
                border: `1px solid ${isActive ? "#2b5cff" : "#e4e7ec"}`,
                borderRadius: 14,
                padding: "0.9rem 1rem",
                display: "grid",
                gridTemplateColumns: "minmax(0, 1fr) auto",
                gap: "1rem",
                alignItems: "center",
                background: isActive ? "#eef4ff" : "#ffffff",
              }}
            >
              <div>
                <div style={{ fontWeight: 700, display: "flex", gap: "0.5rem", alignItems: "center" }}>
                  {isActive && <span title="Активный план">⭐</span>}
                  <span>{plan.name || plan.id}</span>
                </div>
                <div style={{ color: "#475467", fontSize: "0.9rem", marginTop: "0.25rem" }}>
                  Начало: <strong>{formatRu(plan.start_date)}</strong> · Дней: {plan.total_days ?? "?"} · Дорожек: {plan.total_tracks ?? "?"}
                  {plan.created_at && <> · Создан: {plan.created_at}</>}
                </div>
              </div>
              <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
                {!isActive && (
                  <Button
                    variant="secondary"
                    onClick={() => activateMutation.mutate(plan.id)}
                    disabled={activateMutation.isPending}
                  >
                    Сделать активным
                  </Button>
                )}
                {onOpenPlanCalendar && (
                  <Button variant="ghost" onClick={() => onOpenPlanCalendar(plan.id)}>
                    Открыть календарь
                  </Button>
                )}
                <Button
                  variant="danger"
                  onClick={() => setConfirmDelete(plan)}
                  disabled={deleteMutation.isPending}
                >
                  Удалить
                </Button>
              </div>
            </div>
          );
        })}
      </div>

      <Modal
        open={confirmDelete !== null}
        onClose={() => setConfirmDelete(null)}
        title="Удалить план?"
      >
        {confirmDelete && (
          <div style={{ display: "grid", gap: "1rem" }}>
            <p style={{ margin: 0 }}>
              План <strong>{confirmDelete.name || confirmDelete.id}</strong> будет удалён,
              а связанные плиты возвращены в статус <em>«в производстве»</em>.
            </p>
            {deleteMutation.isError &&
              !isPlanVersionConflict(deleteMutation.error) && (
              <Alert tone="error">Не удалось удалить план.</Alert>
            )}
            <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem" }}>
              <Button variant="ghost" onClick={() => setConfirmDelete(null)}>
                Отмена
              </Button>
              <Button
                variant="danger"
                onClick={() =>
                  deleteMutation.mutate(confirmDelete.id, {
                    onSuccess: () => setConfirmDelete(null),
                  })
                }
                disabled={deleteMutation.isPending}
              >
                {deleteMutation.isPending ? "Удаление…" : "Удалить план"}
              </Button>
            </div>
          </div>
        )}
      </Modal>
    </Card>
  );
};
