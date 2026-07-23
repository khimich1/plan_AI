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
import type { BasketDayKind } from "@/features/production/lib/basketDayKind";
import type {
  DayInfo,
  FillTargetItem,
} from "@/features/production/types/production";

const startOfMonth = (d: Date) => new Date(d.getFullYear(), d.getMonth(), 1);

export type GlobalCalendarViewProps = {
  basket: FillTargetItem[];
  basketKind: BasketDayKind | null;
  basketError: string | null;
  /** daysInfo с родителя (для согласованности валидации); иначе из query. */
  daysInfo?: Record<string, DayInfo>;
  brushTracks: number;
  maxBrushTracks: number;
  freeSlotsByDate?: Record<string, number>;
  onBrushTracksChange: (tracks: number) => void;
  onChipTracksChange: (date: string, tracks: number) => void;
  onDayActivate: (
    iso: string,
    meta: { shiftKey: boolean },
    holidays: Set<string>,
    extraWorkdays: Set<string>,
  ) => void;
  onRemove: (date: string) => void;
  onClear: () => void;
  onProceed: () => void;
  onDismissBasketError?: () => void;
};

export const GlobalCalendarView = ({
  basket,
  basketKind,
  basketError,
  daysInfo: daysInfoProp,
  brushTracks,
  maxBrushTracks,
  freeSlotsByDate,
  onBrushTracksChange,
  onChipTracksChange,
  onDayActivate,
  onRemove,
  onClear,
  onProceed,
  onDismissBasketError,
}: GlobalCalendarViewProps) => {
  const [month, setMonth] = useState(() => startOfMonth(new Date()));
  const [selectedDate, setSelectedDate] = useState<string | null>(null);

  const calendarQuery = useGlobalCalendarQuery();
  const workCalendar = useWorkCalendarQuery();

  const daysInfo = daysInfoProp ?? calendarQuery.data?.days_info ?? {};
  const maxPerDay = Object.values(daysInfo)[0]?.max ?? maxBrushTracks;
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

  const basketTracksByDate = useMemo(() => {
    const out: Record<string, number> = {};
    for (const item of basket) {
      out[item.date] = item.tracks;
    }
    return out;
  }, [basket]);

  const isLoading = calendarQuery.isLoading || workCalendar.isLoading;
  const selectedInfo: DayInfo | undefined = selectedDate
    ? (daysInfo[selectedDate] ?? {
        occupied: 0,
        max: maxPerDay,
        completed: false,
        day_number: 0,
      })
    : undefined;

  return (
    <Card
      title="Календарный план"
      subtitle="Задайте N дорожек, кликните дни или Shift+клик для диапазона. Двойной клик или «i» — подробности дня. Свободные и частично занятые дни нельзя смешивать."
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
          onDayActivate={(iso, meta) => {
            onDismissBasketError?.();
            onDayActivate(iso, meta, holidays, extraWorkdays);
          }}
          onOpenDay={(iso) => {
            onDismissBasketError?.();
            setSelectedDate(iso);
          }}
          highlightedDates={highlightedDates}
          basketTracksByDate={basketTracksByDate}
        />
      )}

      <FillBasket
        items={basket}
        basketKind={basketKind}
        basketError={basketError}
        brushTracks={brushTracks}
        maxBrushTracks={maxBrushTracks}
        freeSlotsByDate={freeSlotsByDate}
        onBrushTracksChange={onBrushTracksChange}
        onChipTracksChange={onChipTracksChange}
        onRemove={onRemove}
        onClear={onClear}
        onProceed={onProceed}
      />

      <DayDrawer
        date={selectedDate}
        summary={selectedInfo}
        basketKind={basketKind}
        onClose={() => setSelectedDate(null)}
      />
    </Card>
  );
};
