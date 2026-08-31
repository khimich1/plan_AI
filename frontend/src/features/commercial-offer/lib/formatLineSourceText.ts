/** Prefill for the «как в списке» field from an order_data row. */
export const formatLineSourceText = (item: Record<string, unknown>): string => {
  const qty = Number(item.qty);
  const qtyPart = Number.isFinite(qty) && qty > 0 ? String(qty) : "";
  const mark = String(item.mark ?? "").trim();
  const name = String(item.name ?? "")
    .trim()
    .replace(/^Плиты\s+/i, "");
  const grade = String(item.concrete_grade ?? "").trim();
  const label = mark || name;
  const productType = String(item.product_type ?? "").trim();
  const withGrade =
    Boolean(grade) &&
    productType !== "plates" &&
    productType !== "steps";
  const parts = withGrade ? [label, grade, qtyPart] : [label, qtyPart];
  return parts.filter((part) => part.length > 0).join(" ");
};
