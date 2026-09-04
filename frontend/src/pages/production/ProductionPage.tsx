import { useCallback, useEffect, useMemo, useState } from "react";
import { useSearchParams } from "react-router";
import { Alert } from "@/shared/ui/Alert";
import { ProductionTabs } from "@/features/production/components/ProductionTabs";
import { GlobalCalendarView } from "@/features/production/components/GlobalCalendarView";
import { CreatePlanWizard } from "@/features/production/components/CreatePlanWizard";
import { PlansList } from "@/features/production/components/PlansList";
import { WorkCalendarEditor } from "@/features/production/components/WorkCalendarEditor";
import { SgpWarehouseView } from "@/features/production/components/SgpWarehouseView";
import { KpInWorkView } from "@/features/production/components/KpInWorkView";
import { useGlobalCalendarQuery } from "@/features/production/hooks/useProductionQueries";
import {
  getBasketKind,
  type BasketDayKind,
} from "@/features/production/lib/basketDayKind";
import {
  datesBetweenInclusive,
  paintDays,
} from "@/features/production/lib/calendarRange";
import type {
  DayInfo,
  FillTargetItem,
  ProductionTab,
} from "@/features/production/types/production";

const VALID_TABS: readonly ProductionTab[] = [
  "calendar",
  "create",
  "plans",
  "in-work",
  "work-calendar",
  "sgp",
];

const CALENDAR_HINT = "Сначала выберите дни на календаре.";

const parseTab = (value: string | null): ProductionTab => {
  if (value && (VALID_TABS as readonly string[]).includes(value)) {
    return value as ProductionTab;
  }
  return "calendar";
};

const emptyDayInfo = (max = 5): DayInfo => ({
  occupied: 0,
  max,
  completed: false,
  day_number: 0,
});

const mergePainted = (
  prev: FillTargetItem[],
  painted: FillTargetItem[],
): FillTargetItem[] => {
  const byDate = new Map(prev.map((item) => [item.date, item]));
  for (const item of painted) {
    byDate.set(item.date, item);
  }
  return [...byDate.values()].sort((a, b) => a.date.localeCompare(b.date));
};

export const ProductionPage = () => {
  const [searchParams, setSearchParams] = useSearchParams();
  const tab = parseTab(searchParams.get("tab"));

  // Корзина дозаполнения живёт на странице, чтобы пережить переход
  // между вкладками («календарь» → «дозаполнение»).
  const [basket, setBasket] = useState<FillTargetItem[]>([]);
  // fillRequest — одноразовое сообщение для CreatePlanWizard. Когда он
  // подхватил его, он зовёт onFillRequestConsumed, и мы чистим состояние.
  const [fillRequest, setFillRequest] = useState<FillTargetItem[] | null>(null);
  const [basketError, setBasketError] = useState<string | null>(null);
  const [calendarHint, setCalendarHint] = useState<string | null>(null);
  const [brushTracks, setBrushTracks] = useState<number | null>(null);
  const [selectionAnchor, setSelectionAnchor] = useState<string | null>(null);

  const calendarQuery = useGlobalCalendarQuery();
  const daysInfo = calendarQuery.data?.days_info ?? {};
  const maxPerDay = Object.values(daysInfo)[0]?.max ?? 5;
  const resolvedBrushTracks = Math.max(
    1,
    Math.min(brushTracks ?? maxPerDay, maxPerDay),
  );

  const basketKind: BasketDayKind | null = useMemo(
    () => getBasketKind(basket, daysInfo),
    [basket, daysInfo],
  );

  const freeSlotsByDate = useMemo(() => {
    const out: Record<string, number> = {};
    for (const item of basket) {
      const info = daysInfo[item.date] ?? emptyDayInfo(maxPerDay);
      out[item.date] = Math.max(0, info.max - info.occupied);
    }
    return out;
  }, [basket, daysInfo, maxPerDay]);

  useEffect(() => {
    if (!searchParams.get("tab")) {
      setSearchParams({ tab: "calendar" }, { replace: true });
    }
  }, [searchParams, setSearchParams]);

  // ?tab=create без корзины / fillRequest → календарь + подсказка
  useEffect(() => {
    if (tab !== "create") return;
    if (fillRequest && fillRequest.length > 0) return;
    if (basket.length > 0) return;
    setCalendarHint(CALENDAR_HINT);
    setSearchParams({ tab: "calendar" }, { replace: true });
  }, [tab, fillRequest, basket.length, setSearchParams]);

  const onTabChange = (next: ProductionTab) => {
    setSearchParams({ tab: next });
  };

  const paintDaysWithCalendar = useCallback(
    (
      dates: string[],
      holidays: Set<string>,
      extraWorkdays: Set<string>,
    ) => {
      const currentKind = getBasketKind(basket, daysInfo);
      const { added, error } = paintDays({
        dates,
        brushTracks: resolvedBrushTracks,
        daysInfo,
        basketKind: currentKind,
        holidays,
        extraWorkdays,
        defaultMax: maxPerDay,
      });
      if (added.length > 0) {
        setBasket((prev) => mergePainted(prev, added));
        setCalendarHint(null);
      }
      setBasketError(error);
      return added;
    },
    [basket, resolvedBrushTracks, daysInfo, maxPerDay],
  );

  const handleDayActivate = useCallback(
    (
      iso: string,
      meta: { shiftKey: boolean },
      holidays: Set<string>,
      extraWorkdays: Set<string>,
    ) => {
      const inBasket = basket.some((item) => item.date === iso);

      if (meta.shiftKey && selectionAnchor) {
        const range = datesBetweenInclusive(selectionAnchor, iso);
        paintDaysWithCalendar(range, holidays, extraWorkdays);
        setSelectionAnchor(iso);
        return;
      }

      // Plain click (or Shift without anchor): toggle
      if (inBasket && !meta.shiftKey) {
        setBasket((prev) => prev.filter((p) => p.date !== iso));
        setBasketError(null);
        setSelectionAnchor(iso);
        return;
      }

      paintDaysWithCalendar([iso], holidays, extraWorkdays);
      setSelectionAnchor(iso);
    },
    [basket, paintDaysWithCalendar, selectionAnchor],
  );

  const updateChipTracks = useCallback(
    (date: string, tracks: number) => {
      const info = daysInfo[date] ?? emptyDayInfo(maxPerDay);
      const freeSlots = Math.max(1, info.max - info.occupied);
      const safe = Math.max(1, Math.min(tracks, freeSlots));
      setBasket((prev) =>
        prev.map((item) =>
          item.date === date ? { ...item, tracks: safe } : item,
        ),
      );
    },
    [daysInfo, maxPerDay],
  );

  const removeFromBasket = useCallback((date: string) => {
    setBasket((prev) => prev.filter((p) => p.date !== date));
    setBasketError(null);
  }, []);

  const clearBasket = useCallback(() => {
    setBasket([]);
    setBasketError(null);
    setSelectionAnchor(null);
  }, []);

  const handleProceed = useCallback(() => {
    if (basket.length === 0) return;
    setFillRequest(basket);
    setSearchParams({ tab: "create" });
  }, [basket, setSearchParams]);

  const handleFillConsumed = useCallback(() => {
    // Только fillRequest: корзину чистим на cancel/success.
    // Иначе redirect `?tab=create` без корзины сработает сразу после consume.
    setFillRequest(null);
  }, []);

  return (
    <main style={{ maxWidth: 1280, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <header>
          <h1 style={{ margin: 0, fontSize: "1.75rem" }}>Планирование производства плит</h1>
          <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
            Сводный календарь загрузки, создание и управление производственными планами.
          </p>
        </header>

        <ProductionTabs value={tab} onChange={onTabChange} />

        {calendarHint && tab === "calendar" && (
          <Alert tone="info">{calendarHint}</Alert>
        )}

        {tab === "calendar" && (
          <GlobalCalendarView
            basket={basket}
            basketKind={basketKind}
            basketError={basketError}
            daysInfo={daysInfo}
            brushTracks={resolvedBrushTracks}
            maxBrushTracks={maxPerDay}
            freeSlotsByDate={freeSlotsByDate}
            onBrushTracksChange={setBrushTracks}
            onChipTracksChange={updateChipTracks}
            onDayActivate={handleDayActivate}
            onRemove={removeFromBasket}
            onClear={clearBasket}
            onProceed={handleProceed}
            onDismissBasketError={() => setBasketError(null)}
          />
        )}
        {tab === "create" && (
          <CreatePlanWizard
            onCreated={() => {
              setBasket([]);
              setBasketError(null);
              setSelectionAnchor(null);
              setSearchParams({ tab: "calendar" });
            }}
            fillRequest={fillRequest}
            onFillRequestConsumed={handleFillConsumed}
            onCancelFill={() => {
              setFillRequest(null);
              setBasket([]);
              setBasketError(null);
              setSelectionAnchor(null);
              setSearchParams({ tab: "calendar" });
            }}
          />
        )}
        {tab === "plans" && (
          <PlansList onOpenPlanCalendar={() => setSearchParams({ tab: "calendar" })} />
        )}
        {tab === "in-work" && <KpInWorkView />}
        {tab === "sgp" && <SgpWarehouseView />}
        {tab === "work-calendar" && <WorkCalendarEditor />}
      </div>
    </main>
  );
};
