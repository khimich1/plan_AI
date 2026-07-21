export const formatMoney = (value: number | null | undefined): string => {
  const num = Number(value ?? 0);
  if (!Number.isFinite(num)) {
    return "0 ₽";
  }
  return `${num.toLocaleString("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} ₽`;
};

export const truncate = (value: string | null | undefined, max = 28): string => {
  const str = (value ?? "").toString();
  if (str.length <= max) {
    return str;
  }
  return `${str.slice(0, Math.max(0, max - 1))}…`;
};

export const statusEmoji = (status: string | null | undefined): string => {
  switch (status) {
    case "в архиве":
      return "📦";
    case "в работе":
      return "🏭";
    case "выполнено":
      return "✅";
    case "отклонено":
      return "❌";
    default:
      return "❓";
  }
};
