import type {
  PileCostInputsBody,
  PileCostOverheadInput,
  PileCostRecipeInput,
} from "@/features/price-desk/types/pileCost";

const ARTICLE_LABELS: Record<string, string> = {
  cement: "Цемент",
  crushed_20: "Щебень 20",
  crushed_5: "Щебень (Щ.5)",
  sand: "Песок",
  additive: "Добавка",
  wire_5: "Проволока Ø5",
  ai_10: "А-I Ø10",
  a500c_10: "А500С Ø10",
  a500c_12: "А500С Ø12",
};

const GRADE_LABELS: Record<string, string> = {
  B15: "В15",
  B20: "В20",
  B22_5: "В22,5",
  B25: "В25",
  B30_granite: "В30, гранит",
};

export type RecipeField = "cement_kg_m3" | "sand_kg_m3" | "crushed_kg_m3";
export type OverheadField = keyof PileCostOverheadInput;

export type KnobDraft = {
  prices: Record<string, string>;
  recipes: Record<string, Record<RecipeField, string>>;
  overhead: Record<OverheadField, string>;
};

export const articleLabel = (code: string): string => ARTICLE_LABELS[code] ?? code;

export const gradeLabel = (code: string): string => GRADE_LABELS[code] ?? code;

/** Полная дробная часть, без округления до рубля и без группировки разрядов. */
export const formatPileAmount = (value: number | null | undefined): string => {
  if (value == null || !Number.isFinite(value)) {
    return "—";
  }
  const text = value.toLocaleString("en-US", {
    maximumFractionDigits: 12,
    minimumFractionDigits: 0,
    useGrouping: false,
  });
  return text.replace(".", ",");
};

export const draftFromInputs = (inputs: PileCostInputsBody): KnobDraft => ({
  prices: Object.fromEntries(inputs.prices.map((row) => [row.article, String(row.price)])),
  recipes: Object.fromEntries(
    inputs.recipes.map((row) => [
      row.concrete_grade,
      {
        cement_kg_m3: String(row.cement_kg_m3),
        sand_kg_m3: String(row.sand_kg_m3),
        crushed_kg_m3: String(row.crushed_kg_m3),
      },
    ]),
  ),
  overhead: {
    general: String(inputs.overhead.general),
    payroll: String(inputs.overhead.payroll),
    forming: String(inputs.overhead.forming),
    other: String(inputs.overhead.other),
  },
});

const parseDecimal = (raw: string): number | null => {
  const text = raw.trim().replace(/\s/g, "").replace(",", ".");
  if (!/^-?\d+(\.\d+)?$/.test(text)) {
    return null;
  }
  const value = Number(text);
  return Number.isFinite(value) ? value : null;
};

const priceRow = (
  article: string,
  raw: string,
): { article: string; price: number } | { error: string } => {
  const value = parseDecimal(raw);
  const label = articleLabel(article);
  if (value == null) {
    return { error: `Цена «${label}» должна быть числом` };
  }
  if (value <= 0) {
    return { error: `Цена «${label}» должна быть больше нуля` };
  }
  return { article, price: value };
};

const recipeRow = (
  grade: string,
  raw: Record<RecipeField, string> | undefined,
): PileCostRecipeInput | { error: string } => {
  const label = gradeLabel(grade);
  const cement = parseDecimal(raw?.cement_kg_m3 ?? "");
  const sand = parseDecimal(raw?.sand_kg_m3 ?? "");
  const crushed = parseDecimal(raw?.crushed_kg_m3 ?? "");
  if (cement == null || sand == null || crushed == null) {
    return { error: `Рецепт ${label} должен быть числом` };
  }
  if (cement < 0 || sand < 0 || crushed < 0) {
    return { error: `Рецепт ${label} не может быть отрицательным` };
  }
  return {
    concrete_grade: grade,
    cement_kg_m3: cement,
    sand_kg_m3: sand,
    crushed_kg_m3: crushed,
  };
};

const overheadValue = (field: OverheadField, raw: string): number | { error: string } => {
  const names: Record<OverheadField, string> = {
    general: "Общие производственные",
    payroll: "ФОТ",
    forming: "Обслуживание формовки",
    other: "Обслуживание, остальное",
  };
  const value = parseDecimal(raw);
  if (value == null) {
    return { error: `${names[field]} должно быть числом` };
  }
  if (value < 0) {
    return { error: `${names[field]} не может быть отрицательным` };
  }
  return value;
};

/**
 * Тело PUT: только ручки, в порядке ответа сервера.
 * Объём сваи и килограммы арматуры сюда не входят.
 */
export const buildInputs = (
  draft: KnobDraft,
  source: PileCostInputsBody,
): { ok: true; body: PileCostInputsBody } | { ok: false; error: string } => {
  const prices = [];
  for (const row of source.prices) {
    const built = priceRow(row.article, draft.prices[row.article] ?? "");
    if ("error" in built) {
      return { ok: false, error: built.error };
    }
    prices.push(built);
  }

  const recipes = [];
  for (const row of source.recipes) {
    const built = recipeRow(row.concrete_grade, draft.recipes[row.concrete_grade]);
    if ("error" in built) {
      return { ok: false, error: built.error };
    }
    recipes.push(built);
  }

  const overhead: Partial<PileCostOverheadInput> = {};
  for (const field of ["general", "payroll", "forming", "other"] as const) {
    const built = overheadValue(field, draft.overhead[field] ?? "");
    if (typeof built !== "number") {
      return { ok: false, error: built.error };
    }
    overhead[field] = built;
  }

  return {
    ok: true,
    body: {
      prices,
      recipes,
      overhead: overhead as PileCostOverheadInput,
    },
  };
};
