import { useState, type ReactNode } from "react";
import type { UseMutationResult } from "@tanstack/react-query";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import { getErrorMessage } from "@/shared/lib/apiError";
import { ResetConfirmDialog } from "@/features/admin/components/ResetConfirmDialog";
import {
  useCalendarResetMutation,
  useDbStatsQuery,
  useFullResetMutation,
  useKpResetMutation,
  usePlansResetMutation,
  useRecoverPlatesMutation,
} from "@/features/admin/hooks/useAdminQueries";
import type {
  DbResetReport,
  DbStatsResponse,
  RecoverPlatesResponse,
  ResetVariant,
} from "@/features/admin/types/admin";

type Props = {
  open: boolean;
  onClose: () => void;
};

type DialogConfig = {
  variant: ResetVariant;
  title: string;
  description: ReactNode;
  confirmLabel: string;
  confirmKeyword?: string;
  mutation: UseMutationResult<DbResetReport, Error, void>;
};

const StatRow = ({ label, value }: { label: string; value: ReactNode }) => (
  <div
    style={{
      display: "flex",
      justifyContent: "space-between",
      gap: "1rem",
      padding: "0.4rem 0",
      borderBottom: "1px dashed #eef2ff",
      fontSize: "0.95rem",
    }}
  >
    <span style={{ color: "#475467" }}>{label}</span>
    <strong style={{ color: "#23366f" }}>{value}</strong>
  </div>
);

const StatsBlock = ({ stats }: { stats: DbStatsResponse }) => (
  <div>
    <StatRow label="Всего КП" value={stats.kp_total} />
    <StatRow label="КП в работе" value={stats.kp_in_work} />
    <StatRow label="КП выполнено" value={stats.kp_completed} />
    <StatRow label="Плит в работе" value={stats.plates_in_work} />
    <StatRow label="Выполненных плит" value={stats.plates_completed} />
    <StatRow label="Остатков от резки" value={stats.plate_rests} />
    <StatRow label="Файлов планов (JSON)" value={stats.plans_count} />
    <StatRow
      label="current_plan.json"
      value={stats.current_plan_present ? "есть" : "нет"}
    />
  </div>
);

const RecoverResultLine = ({
  result,
}: {
  result: RecoverPlatesResponse | undefined;
}) => {
  if (!result) {
    return null;
  }
  if (result.recovered_records === 0) {
    return <Alert tone="info">Застрявших плит не найдено.</Alert>;
  }
  return (
    <Alert tone="success">
      Восстановлено записей: <strong>{result.recovered_records}</strong>
    </Alert>
  );
};

export const DbManagementModal = ({ open, onClose }: Props) => {
  const statsQuery = useDbStatsQuery(open);

  const fullReset = useFullResetMutation();
  const kpReset = useKpResetMutation();
  const plansReset = usePlansResetMutation();
  const calendarReset = useCalendarResetMutation();
  const recoverPlates = useRecoverPlatesMutation();

  const [dialog, setDialog] = useState<DialogConfig | null>(null);

  const handleClose = () => {
    fullReset.reset();
    kpReset.reset();
    plansReset.reset();
    calendarReset.reset();
    recoverPlates.reset();
    setDialog(null);
    onClose();
  };

  const openDialog = (config: DialogConfig) => {
    config.mutation.reset();
    setDialog(config);
  };

  const closeDialog = () => {
    if (dialog?.mutation.isPending) {
      return;
    }
    setDialog(null);
  };

  const onConfirmDialog = async () => {
    if (!dialog) {
      return;
    }
    try {
      await dialog.mutation.mutateAsync();
      setDialog(null);
      statsQuery.refetch();
    } catch {
      // ошибка показывается через mutation.error
    }
  };

  return (
    <>
      <Modal
        open={open}
        onClose={handleClose}
        title="Управление базой данных"
        maxWidth={640}
      >
        <div style={{ display: "grid", gap: "1.25rem" }}>
          <Alert tone="warning">
            Действия в этом окне необратимы. Перед массовым обнулением
            рекомендуется остановить Telegram-бот, чтобы избежать гонки за
            файлы планов и базу данных.
          </Alert>

          <section>
            <h3
              style={{
                margin: "0 0 0.5rem",
                fontSize: "1rem",
                color: "#23366f",
              }}
            >
              Состояние БД
            </h3>
            {statsQuery.isLoading && (
              <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
                <Spinner /> <span>Загружаем статистику...</span>
              </div>
            )}
            {statsQuery.isError && (
              <Alert tone="error">{getErrorMessage(statsQuery.error)}</Alert>
            )}
            {statsQuery.data && <StatsBlock stats={statsQuery.data} />}
          </section>

          <section style={{ display: "grid", gap: "0.6rem" }}>
            <h3
              style={{
                margin: "0",
                fontSize: "1rem",
                color: "#b42318",
              }}
            >
              Опасные операции
            </h3>

            <Button
              variant="danger"
              fullWidth
              onClick={() =>
                openDialog({
                  variant: "full",
                  title: "Полное обнуление БД",
                  description: (
                    <span>
                      Будут удалены <strong>все КП, плиты, файлы, остатки,
                      выполненные плиты и журнал статусов</strong>, а также
                      все <strong>JSON-планы производства</strong> и{" "}
                      <strong>производственный календарь</strong>. Учётная запись администратора
                      сохранится. Действие необратимо.
                    </span>
                  ),
                  confirmLabel: "Обнулить ВСЁ",
                  confirmKeyword: "ОБНУЛИТЬ",
                  mutation: fullReset,
                })
              }
            >
              Полное обнуление (БД + планы + календарь)
            </Button>

            <Button
              variant="danger"
              fullWidth
              onClick={() =>
                openDialog({
                  variant: "kp-only",
                  title: "Обнуление только КП",
                  description: (
                    <span>
                      Будут удалены таблицы КП: <code>KP_offers</code>,{" "}
                      <code>kp_plates</code>, <code>kp_files</code>,{" "}
                      <code>kp_meta</code>. Записи о выполненных плитах и
                      остатках напрямую не очищаются (но удаляются каскадом
                      при удалении родительских КП).
                    </span>
                  ),
                  confirmLabel: "Обнулить КП",
                  mutation: kpReset,
                })
              }
            >
              Только КП (без выполненных плит и остатков)
            </Button>

            <Button
              variant="danger"
              fullWidth
              onClick={() =>
                openDialog({
                  variant: "plans-only",
                  title: "Удалить JSON-планы производства",
                  description: (
                    <span>
                      Будут удалены файлы <code>data/plans/*.json</code>,{" "}
                      <code>data/plans_metadata.json</code> и{" "}
                      <code>data/current_plan.json</code>. Содержимое SQLite-базы и
                      календарь не пострадают.
                    </span>
                  ),
                  confirmLabel: "Удалить планы",
                  mutation: plansReset,
                })
              }
            >
              Только планы (JSON)
            </Button>

            <Button
              variant="danger"
              fullWidth
              onClick={() =>
                openDialog({
                  variant: "calendar-only",
                  title: "Сбросить производственный календарь",
                  description: (
                    <span>
                      Файл <code>work_calendar.json</code> будет приведён к
                      пустому состоянию (без дополнительных праздников и
                      рабочих дней).
                    </span>
                  ),
                  confirmLabel: "Сбросить календарь",
                  mutation: calendarReset,
                })
              }
            >
              Только календарь
            </Button>
          </section>

          <section style={{ display: "grid", gap: "0.6rem" }}>
            <h3
              style={{
                margin: 0,
                fontSize: "1rem",
                color: "#23366f",
              }}
            >
              Сервисные операции
            </h3>

            <Button
              variant="secondary"
              fullWidth
              onClick={() => recoverPlates.mutate()}
              disabled={recoverPlates.isPending}
            >
              {recoverPlates.isPending
                ? "Восстанавливаем..."
                : "Восстановить застрявшие плиты"}
            </Button>
            {recoverPlates.isError && (
              <Alert tone="error">{getErrorMessage(recoverPlates.error)}</Alert>
            )}
            {recoverPlates.isSuccess && (
              <RecoverResultLine result={recoverPlates.data} />
            )}
          </section>

          <div style={{ display: "flex", justifyContent: "flex-end" }}>
            <Button variant="ghost" onClick={handleClose}>
              Закрыть
            </Button>
          </div>
        </div>
      </Modal>

      {dialog && (
        <ResetConfirmDialog
          open={dialog !== null}
          onClose={closeDialog}
          title={dialog.title}
          description={dialog.description}
          confirmLabel={dialog.confirmLabel}
          confirmKeyword={dialog.confirmKeyword}
          isPending={dialog.mutation.isPending}
          isError={dialog.mutation.isError}
          error={dialog.mutation.error}
          onConfirm={onConfirmDialog}
        />
      )}
    </>
  );
};
