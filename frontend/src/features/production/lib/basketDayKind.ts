import type { DayInfo, FillTargetItem } from "@/features/production/types/production";

/** Тип дня по занятости: свободный / частично занятый / полный. */
export type DayKind = "empty" | "partial" | "full";

/** Тип дней в корзине — смешивать empty и partial нельзя. */
export type BasketDayKind = "empty" | "partial";

export function getDayKind(info: Pick<DayInfo, "occupied" | "max">): DayKind {
  const freeSlots = Math.max(0, info.max - info.occupied);
  if (info.occupied <= 0) return "empty";
  if (freeSlots <= 0) return "full";
  return "partial";
}

export function getBasketKind(
  items: FillTargetItem[],
  daysInfo: Record<string, DayInfo>,
): BasketDayKind | null {
  if (items.length === 0) return null;
  const first = daysInfo[items[0].date];
  // День вне диапазона days_info = свободный (ещё не в планах).
  if (!first) return "empty";
  const kind = getDayKind(first);
  if (kind === "full") return null;
  return kind;
}

export function canAddDayToBasket(
  basketKind: BasketDayKind | null,
  dayKind: DayKind,
): boolean {
  if (dayKind === "full") return false;
  if (basketKind === null) return true;
  return basketKind === dayKind;
}
