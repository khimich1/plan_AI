import { useMemo, useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Card } from "@/shared/ui/Card";
import { Spinner } from "@/shared/ui/Spinner";
import { DayDrawer } from "@/features/production/components/DayDrawer";
import { FillBasket } from "@/features/production/components/FillBasket";
import { MonthCalendarGrid } from "@/features/production/components/MonthCalendarGrid";
import {
  useGlobalCalendarQuery,
  useWorkCalendarQuery,
} from "@/features/production/hooks/useProductionQueries";
import type { FillTargetItem } from "@/features/production/types/production";

const startOfMonth = (d: Date) => new Date(d.getFullYear(), d.getMonth(), 1);

export type GlobalCalendarViewProps = {
  basket: FillTargetItem[];
  onAdd: (date: string, tracks: number) => void;
  onRemove: (date: string) => void;
  onClear: () => void;
  onProceed: () => void;
};

export const GlobalCalendarView = ({
  basket,
  onAdd,
  onRemove,
  onClear,
  onProceed,
}: GlobalCalendarViewProps) => {
  const [month, setMonth] = useState(() => startOfMonth(new Date()));
  const [selectedDate, setSelectedDate] = useState<string | null>(null);

  const calendarQuery = useGlobalCalendarQuery();
  const workCalendar = useWorkCalendarQuery();

  const daysInfo = calendarQuery.data?.days_info ?? {};
  const holidays = useMemo(
    () => new Set(workCalendar.data?.extra_holidays ?? []),
    [workCalendar.data],
  );
  const extraWorkdays = useMemo(
    () => new Set(workCalendar.data?.extra_workdays ?? []),
    [workCalendar.data],
  );

  const highlightedDates = useMemo(
    () => new Set(basket.map((item) => item.date)),
    [basket],
  );

  const isLoading = calendarQuery.isLoading || workCalendar.isLoading;
  const selectedInfo = selectedDate ? daysInfo[selectedDate] : undefined;
  const alreadyInBasketTracks = selectedDate
    ? basket.find((item) => item.date === selectedDate)?.tracks
    : undefined;

  return (
    <Card
      title="Календарный план"
      subtitle="Сводная загрузка всех планов по датам. Клик по дню — посмотреть содержимое или добавить день в дозаполнение."
    >
      {isLoading && (
        <div style={{ display: "flex", alignItems: "center", gap: "0.5rem" }}>
          <Spinner /> Загрузка календаря…
        </div>
      )}

      {calendarQuery.isError && (
        <Alert tone="error">Не удалось загрузить календарь производства.</Alert>
      )}

      {!isLoading && (
        <MonthCalendarGrid
          daysInfo={daysInfo}
          holidays={holidays}
          extraWorkdays={extraWorkdays}
          month={month}
          onMonthChange={setMonth}
          selectedDate={selectedDate}
          onSelectDate={setSelectedDate}
          highlightedDates={highlightedDates}
        />
      )}

      <FillBasket
        items={basket}
        onRemove={onRemove}
        onClear={onClear}
        onProceed={onProceed}
      />

      <DayDrawer
        date={selectedDate}
        summary={selectedInfo}
        onClose={() => setSelectedDate(null)}
        onAddToFillBasket={onAdd}
        alreadyInBasketTracks={alreadyInBasketTracks}
      />
    </Card>
  );
};
