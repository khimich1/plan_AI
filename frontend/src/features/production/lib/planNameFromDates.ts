/** Автоимя плана по отсортированным датам корзины (ISO YYYY-MM-DD). */
export function planNameFromDates(dates: string[]): string {
  if (dates.length === 0) return "";
  const sorted = [...dates].sort();
  const day = (iso: string) => iso.split("-")[2];
  const dayMonth = (iso: string) => {
    const [, m, d] = iso.split("-");
    return `${d}.${m}`;
  };
  if (sorted.length === 1) return `План ${dayMonth(sorted[0])}`;
  // «План 23–25.07» — день начала без месяца, конец с DD.MM
  return `План ${day(sorted[0])}–${dayMonth(sorted[sorted.length - 1])}`;
}
