export const PRESET_TRACKS = [1, 2, 3, 4, 5];
export const MAX_PER_DAY = 5;

export const todayISO = () => {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
};

export const parseISODate = (iso: string) => {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
};

export const startOfMonth = (d: Date) => new Date(d.getFullYear(), d.getMonth(), 1);

export const formatRu = (iso: string) => {
  const d = parseISODate(iso);
  return d.toLocaleDateString("ru-RU", { day: "2-digit", month: "2-digit" });
};
