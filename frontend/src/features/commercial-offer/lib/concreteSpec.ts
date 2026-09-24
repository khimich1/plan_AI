export const GOST_FROST_RESISTANCE = [
  "F50",
  "F75",
  "F100",
  "F150",
  "F200",
  "F300",
  "F400",
  "F500",
  "F600",
  "F800",
  "F1000",
] as const;

export const GOST_WATERPROOFNESS = [
  "W2",
  "W4",
  "W6",
  "W8",
  "W10",
  "W12",
  "W14",
  "W16",
  "W18",
  "W20",
] as const;

export type ConcreteSpecFields = {
  frost_resistance: string | null;
  waterproofness: string | null;
  concrete_aggregate: string | null;
  concrete_spec_source: string | null;
};

export type ConcreteSpecPatch =
  | { concrete_spec_source: "table"; concrete_aggregate: "granite" | "ordinary" }
  | { concrete_spec_source: "manual"; frost_resistance: string; waterproofness: string };

export type TableConcretePair = {
  aggregate: "granite" | "ordinary";
  frost_resistance: string;
  waterproofness: string;
  label: string;
};

type TableRow = {
  frost: string;
  granite: string;
  ordinary: string | null;
};

const TABLE: Record<string, TableRow> = {
  "B7.5": { frost: "F50", granite: "W2", ordinary: "W2" },
  "B12.5": { frost: "F50", granite: "W2", ordinary: "W2" },
  B15: { frost: "F100", granite: "W4", ordinary: "W4" },
  B20: { frost: "F100", granite: "W4", ordinary: "W4" },
  "B22.5": { frost: "F200", granite: "W6", ordinary: "W4" },
  B25: { frost: "F200", granite: "W8", ordinary: "W6" },
  B30: { frost: "F300", granite: "W10", ordinary: null },
  M400: { frost: "F300", granite: "W10", ordinary: null },
  B35: { frost: "F300", granite: "W10", ordinary: null },
  M500: { frost: "F300", granite: "W12", ordinary: null },
  B40: { frost: "F300", granite: "W12", ordinary: null },
  B45: { frost: "F300", granite: "W12", ordinary: null },
};

const lookupKey = (concreteGrade: string): string | null => {
  const text = concreteGrade.trim().replace(/М/g, "M").replace(/м/g, "M");
  if (!text) {
    return null;
  }
  let key = text.toUpperCase().replace(/_/g, ".");
  if (key.endsWith(".GRANITE")) {
    key = key.slice(0, -".GRANITE".length);
  }
  return key;
};

export const formatFrostPair = (
  frost: string | null | undefined,
  water: string | null | undefined,
): string => {
  const frostText = (frost ?? "").trim();
  const waterText = (water ?? "").trim();
  if (!frostText || !waterText) {
    return "—";
  }
  return `${frostText} · ${waterText}`;
};

const pair = (
  aggregate: "granite" | "ordinary",
  frost: string,
  water: string,
): TableConcretePair => ({
  aggregate,
  frost_resistance: frost,
  waterproofness: water,
  label: formatFrostPair(frost, water),
});

export const tablePairsForGrade = (concreteGrade: string): TableConcretePair[] => {
  const key = lookupKey(concreteGrade);
  if (!key) {
    return [];
  }
  const row = TABLE[key];
  if (!row) {
    return [];
  }
  const pairs = [pair("granite", row.frost, row.granite)];
  if (row.ordinary && row.ordinary !== row.granite) {
    pairs.push(pair("ordinary", row.frost, row.ordinary));
  }
  return pairs;
};

export const isOffTablePair = (
  concreteGrade: string,
  frost: string | null | undefined,
  water: string | null | undefined,
): boolean => {
  const label = formatFrostPair(frost, water);
  if (label === "—") {
    return false;
  }
  return !tablePairsForGrade(concreteGrade).some((item) => item.label === label);
};

const trimmedField = (item: Record<string, unknown>, key: string): string | null => {
  const value = item[key];
  if (value == null) {
    return null;
  }
  const text = String(value).trim();
  return text || null;
};

export const readConcreteSpec = (item: Record<string, unknown>): ConcreteSpecFields => {
  const source = trimmedField(item, "concrete_spec_source");
  const aggregateValue = item.concrete_aggregate;
  let aggregate: string | null = null;
  if (aggregateValue != null) {
    const raw = String(aggregateValue);
    aggregate = source === "manual" && raw.trim() === "" ? "" : raw.trim() || null;
  }
  return {
    frost_resistance: trimmedField(item, "frost_resistance"),
    waterproofness: trimmedField(item, "waterproofness"),
    concrete_aggregate: aggregate,
    concrete_spec_source: source,
  };
};

export const specFromPreviewRow = (row: Partial<ConcreteSpecFields>): ConcreteSpecFields => ({
  frost_resistance: row.frost_resistance ?? null,
  waterproofness: row.waterproofness ?? null,
  concrete_aggregate: row.concrete_aggregate ?? null,
  concrete_spec_source: row.concrete_spec_source ?? null,
});
