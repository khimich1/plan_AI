import type {
  ProductType,
  WizardNextRequiredAction,
  WizardStepId,
} from "@/features/commercial-offer/types/commercialOffer";

/** URL segment of the per-product draft endpoints (note: bridge_piles → "bridge-piles"). */
export type ProductEndpointSegment = "plates" | "piles" | "steps" | "marches" | "bridge-piles" | "fbs";

/** metadata.* field holding the product's source batches. */
export type DraftBatchesField =
  | "plate_batches"
  | "pile_batches"
  | "step_batches"
  | "march_batches"
  | "bridge_pile_batches"
  | "fbs_batches";

export type ProductTypeConfig = {
  productType: ProductType;
  /** Wizard input step of the product (currently identical to productType). */
  inputStep: WizardStepId;
  endpointSegment: ProductEndpointSegment;
  batchesField: DraftBatchesField;
  /** piles/marches/bridge_piles/fbs expose concrete-grade editing; plates/steps do not. */
  supportsGrades: boolean;
  /** Simple KP flow (no breakdown / plate resolve gates): everything except plates. */
  isSimpleKp: boolean;
  ingestAction: WizardNextRequiredAction;
  labels: {
    /** "Сваи" — nominative plural, product type column in the result table. */
    nounPlural: string;
    /** "свай" — wizard messages and the result readiness line ("N свай в заказе"). */
    nounGenitivePlural: string;
  };
};

export const PRODUCT_TYPE_CONFIG: Record<ProductType, ProductTypeConfig> = {
  plates: {
    productType: "plates",
    inputStep: "plates",
    endpointSegment: "plates",
    batchesField: "plate_batches",
    supportsGrades: false,
    isSimpleKp: false,
    ingestAction: "ingest_plates",
    labels: { nounPlural: "Плиты", nounGenitivePlural: "плит" },
  },
  piles: {
    productType: "piles",
    inputStep: "piles",
    endpointSegment: "piles",
    batchesField: "pile_batches",
    supportsGrades: true,
    isSimpleKp: true,
    ingestAction: "ingest_piles",
    labels: { nounPlural: "Сваи", nounGenitivePlural: "свай" },
  },
  steps: {
    productType: "steps",
    inputStep: "steps",
    endpointSegment: "steps",
    batchesField: "step_batches",
    supportsGrades: false,
    isSimpleKp: true,
    ingestAction: "ingest_steps",
    labels: { nounPlural: "Ступени", nounGenitivePlural: "ступеней" },
  },
  marches: {
    productType: "marches",
    inputStep: "marches",
    endpointSegment: "marches",
    batchesField: "march_batches",
    supportsGrades: true,
    isSimpleKp: true,
    ingestAction: "ingest_marches",
    labels: { nounPlural: "Марши", nounGenitivePlural: "маршей" },
  },
  bridge_piles: {
    productType: "bridge_piles",
    inputStep: "bridge_piles",
    endpointSegment: "bridge-piles",
    batchesField: "bridge_pile_batches",
    supportsGrades: true,
    isSimpleKp: true,
    ingestAction: "ingest_bridge_piles",
    labels: { nounPlural: "Мостовые сваи", nounGenitivePlural: "мостовых свай" },
  },
  fbs: {
    productType: "fbs",
    inputStep: "fbs",
    endpointSegment: "fbs",
    batchesField: "fbs_batches",
    supportsGrades: true,
    isSimpleKp: true,
    ingestAction: "ingest_fbs",
    labels: { nounPlural: "ФБС", nounGenitivePlural: "ФБС" },
  },
};

/** Unknown / missing product types resolve to plates — keeps legacy drafts working. */
export const getProductTypeConfig = (type: ProductType | null | undefined): ProductTypeConfig => {
  if (type && type in PRODUCT_TYPE_CONFIG) {
    return PRODUCT_TYPE_CONFIG[type];
  }
  return PRODUCT_TYPE_CONFIG.plates;
};

export const INGEST_REQUIRED_MESSAGE = "Сначала распознайте и получите хотя бы одну позицию в заказе.";

/**
 * «Нельзя перейти дальше» по серверному next_required_action. Действия без записи
 * (none, select_manager, complete_client_terms, post_calculate) отдают
 * product-specific fallback у вызывающей стороны.
 */
export const NEXT_REQUIRED_ACTION_MESSAGES: Partial<Record<WizardNextRequiredAction, string>> = {
  ingest_plates: INGEST_REQUIRED_MESSAGE,
  ingest_piles: INGEST_REQUIRED_MESSAGE,
  ingest_steps: INGEST_REQUIRED_MESSAGE,
  ingest_marches: INGEST_REQUIRED_MESSAGE,
  ingest_bridge_piles: INGEST_REQUIRED_MESSAGE,
  ingest_fbs: INGEST_REQUIRED_MESSAGE,
  resolve_wide_plates: "Сначала примите решение по позициям шире стандартной.",
  resolve_invalid_widths: "Нестандартная ширина: замените на заводской рез или исключите позицию.",
  resolve_unpriced_plates: "Сначала примите решение по позициям без цены в прайсе.",
};
