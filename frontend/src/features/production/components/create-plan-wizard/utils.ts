export const parseISODate = (iso: string) => {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
};

export const formatRu = (iso: string) => {
  const d = parseISODate(iso);
  return d.toLocaleDateString("ru-RU", { day: "2-digit", month: "2-digit" });
};
