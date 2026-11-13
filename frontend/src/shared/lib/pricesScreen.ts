export type PricesScreen = "queue" | "cost" | "update";

export function pricesQueueHref(): string {
  return "/prices?view=queue";
}

export function pricesCostHref(): string {
  return "/prices?view=cost";
}

export function pricesUpdateHref(): string {
  return "/prices?view=update";
}

/** Нет `view`, `queue` и любое чужое значение — очередь. */
export function pricesScreenFromSearch(search: URLSearchParams): PricesScreen {
  const view = search.get("view");
  if (view === "cost") {
    return "cost";
  }
  if (view === "update") {
    return "update";
  }
  return "queue";
}
