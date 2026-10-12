/** Типы API себестоимости цельных свай — синхронны с app/schemas/pile_cost.py. */

export type PileCostArticlePrice = {
  article: string;
  price: number;
};

/** Рецепт бетона, кг/м³. Это ручки, не килограммы арматуры из нормы. */
export type PileCostRecipeInput = {
  concrete_grade: string;
  cement_kg_m3: number;
  sand_kg_m3: number;
  crushed_kg_m3: number;
};

export type PileCostOverheadInput = {
  general: number;
  payroll: number;
  forming: number;
  other: number;
};

/**
 * Тело PUT. Объёма сваи и килограммов арматуры здесь нет:
 * сервер такие поля отклоняет, клиент их не отправляет.
 */
export type PileCostInputsBody = {
  prices: PileCostArticlePrice[];
  recipes: PileCostRecipeInput[];
  overhead: PileCostOverheadInput;
};

export type PileCostGradePrice = {
  concrete_grade: string;
  buyer_price: number | null;
};

export type PileCostMarkSummary = {
  mark: string;
  section_cm: number;
  grades: PileCostGradePrice[];
};

/** Строка металла или щебня в разложении. Килограммы только для чтения. */
export type PileCostPricedLine = {
  article: string;
  kg: number;
  rub: number;
};

export type PileCostGradeBreakdown = {
  concrete_grade: string;
  materials_rub: number;
  expenses_rub: number;
  cost: number;
  buyer_price: number;
  surcharge_applied: boolean;
  cement_kg: number;
  sand_kg: number;
  crushed_kg: number;
  additive_kg: number;
  cement_rub: number;
  sand_rub: number;
  crushed_rub: number;
  additive_rub: number;
  crushed_article: string;
  lines: PileCostPricedLine[];
};

/** Разложение одной марки. volume_m3 и килограммы строк только для чтения. */
export type PileCostMarkBreakdown = {
  mark: string;
  section_cm: number;
  volume_m3: number;
  grades: PileCostGradeBreakdown[];
};

export type PileCostState = {
  inputs_sha256: string;
  inputs: PileCostInputsBody;
  marks: PileCostMarkSummary[];
  breakdown: PileCostMarkBreakdown | null;
};

export type PileCostSaved = {
  inputs_sha256: string;
  inputs: PileCostInputsBody;
};

export type PileCostDiffExample = {
  key: string;
  old_price: number | null;
  new_price: number | null;
};

export type PileCostDiffExamples = {
  changed: PileCostDiffExample[];
  new: PileCostDiffExample[];
  missing: PileCostDiffExample[];
  unchanged: PileCostDiffExample[];
};

export type PileCostPreview = {
  inputs_sha256: string;
  changed: number;
  new: number;
  missing: number;
  unchanged: number;
  examples: PileCostDiffExamples;
};

export type PileCostApplyRequest = {
  inputs_sha256: string | null;
};
