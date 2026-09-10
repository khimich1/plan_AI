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

export type ProductTypeLabels = {
  /** "Сваи" — nominative plural, product type column in the result table. */
  nounPlural: string;
  /** "свай" — wizard messages and the result readiness line ("N свай в заказе"). */
  nounGenitivePlural: string;
  /** "Шаг 1. Сваи" — input step title. */
  stepTitle: string;
  /** "Список свай" — label of the source text field. */
  listLabel: string;
  /** "Список свай для расчёта" — batch-review editor card title. */
  reviewListTitle: string;
  /** Exact current placeholder, including known copy-paste leftovers. */
  placeholder: string;
  /** "Например: убери строки с B15" — AI instruction placeholder on the source card. */
  aiPlaceholder: string;
  /** "Вставьте текст списка свай или загрузите фото таблицы." */
  emptySubtitle: string;
  /** "Редкий сценарий: опишите, что сделать со списком свай." */
  aiHint: string;
  /** "Добавьте ещё сваи или перейдите к оформлению клиента." — step description with a draft. */
  addMoreDescription: string;
  /** "Загрузите фото или вставьте список свай для расчёта." — step description without a draft. */
  initialDescription: string;
  /** "Изменён список свай — нажмите «Список верен» для пересчёта состава." — preview panel notice. */
  previewChangedMessage: string;
  /** "Список пуст — распознайте сваи." — preview panel empty state. */
  previewEmptyMessage: string;
  /** "Марка, класс бетона, количество и цена — как в документе." — preview panel card subtitle. */
  previewSubtitle: string;
  /** "Не все марки найдены в прайсе — …" — preview panel unpriced alert. */
  previewUnpricedMessage: string;
};

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
  labels: ProductTypeLabels;
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
    labels: {
      nounPlural: "Плиты",
      nounGenitivePlural: "плит",
      stepTitle: "Шаг 1. Плиты",
      listLabel: "Список плит",
      reviewListTitle: "Список плит для расчёта",
      placeholder: "ПБ 78-12-8п 2\n71-12-8 3\nПБ 66-12-8п 4",
      aiPlaceholder: "Например: убери строки с 6п",
      emptySubtitle: "Вставьте текст списка плит или загрузите фото таблицы.",
      aiHint: "Редкий сценарий: опишите, что сделать со списком плит.",
      addMoreDescription: "Добавьте ещё плиты или перейдите к оформлению клиента.",
      initialDescription: "Загрузите фото или вставьте список плит для расчёта.",
      previewChangedMessage: "Изменён список плит — нажмите «Список верен» для пересчёта состава.",
      previewEmptyMessage: "Список пуст — распознайте плиты.",
      previewSubtitle: "Наименование, количество и цена — как в документе. Скидка и доставка учитываются позже.",
      previewUnpricedMessage: "Не все плиты найдены в прайсе — исправьте список перед переходом к клиенту.",
    },
  },
  piles: {
    productType: "piles",
    inputStep: "piles",
    endpointSegment: "piles",
    batchesField: "pile_batches",
    supportsGrades: true,
    isSimpleKp: true,
    ingestAction: "ingest_piles",
    labels: {
      nounPlural: "Сваи",
      nounGenitivePlural: "свай",
      stepTitle: "Шаг 1. Сваи",
      listLabel: "Список свай",
      reviewListTitle: "Список свай для расчёта",
      placeholder: "С120.35-12 B25 5\nС120.35-13и 3",
      aiPlaceholder: "Например: убери строки с B15",
      emptySubtitle: "Вставьте текст списка свай или загрузите фото таблицы.",
      aiHint: "Редкий сценарий: опишите, что сделать со списком свай.",
      addMoreDescription: "Добавьте ещё сваи или перейдите к оформлению клиента.",
      initialDescription: "Загрузите фото или вставьте список свай для расчёта.",
      previewChangedMessage: "Изменён список свай — нажмите «Список верен» для пересчёта состава.",
      previewEmptyMessage: "Список пуст — распознайте сваи.",
      previewSubtitle: "Марка, класс бетона, количество и цена — как в документе.",
      previewUnpricedMessage:
        "Не все марки найдены в прайсе — исправьте список или класс бетона перед переходом к клиенту.",
    },
  },
  steps: {
    productType: "steps",
    inputStep: "steps",
    endpointSegment: "steps",
    batchesField: "step_batches",
    supportsGrades: false,
    isSimpleKp: true,
    ingestAction: "ingest_steps",
    labels: {
      nounPlural: "Ступени",
      nounGenitivePlural: "ступеней",
      stepTitle: "Шаг 1. Ступени",
      listLabel: "Список ступеней",
      reviewListTitle: "Список ступеней для расчёта",
      placeholder: "ЛС11 10\nЛС14-1лев 5\nЛС11-Б-1 2",
      aiPlaceholder: "Например: убери строки с ЛС11",
      emptySubtitle: "Вставьте текст списка ступеней или загрузите фото таблицы.",
      aiHint: "Редкий сценарий: опишите, что сделать со списком ступеней.",
      addMoreDescription: "Добавьте ещё ступени или перейдите к оформлению клиента.",
      initialDescription: "Загрузите фото или вставьте список ступеней для расчёта.",
      previewChangedMessage: "Изменён список ступеней — нажмите «Список верен» для пересчёта состава.",
      previewEmptyMessage: "Список пуст — распознайте ступени.",
      previewSubtitle: "Марка, количество и цена — как в документе.",
      previewUnpricedMessage: "Не все марки найдены в прайсе — исправьте список перед переходом к клиенту.",
    },
  },
  marches: {
    productType: "marches",
    inputStep: "marches",
    endpointSegment: "marches",
    batchesField: "march_batches",
    supportsGrades: true,
    isSimpleKp: true,
    ingestAction: "ingest_marches",
    labels: {
      nounPlural: "Марши",
      nounGenitivePlural: "маршей",
      stepTitle: "Шаг 1. Марши",
      listLabel: "Список маршей",
      reviewListTitle: "Список маршей для расчёта",
      placeholder: "1ЛМ 27-11-14-4 B25 5\nЛМ 2,8 3",
      aiPlaceholder: "Например: убери строки с B15",
      emptySubtitle: "Вставьте текст списка маршей или загрузите фото таблицы.",
      aiHint: "Редкий сценарий: опишите, что сделать со списком маршей.",
      addMoreDescription: "Добавьте ещё марши или перейдите к оформлению клиента.",
      initialDescription: "Загрузите фото или вставьте список маршей для расчёта.",
      previewChangedMessage: "Изменён список маршей — нажмите «Список верен» для пересчёта состава.",
      previewEmptyMessage: "Список пуст — распознайте марши.",
      previewSubtitle: "Марка, класс бетона, количество и цена — как в документе.",
      previewUnpricedMessage:
        "Не все марки найдены в прайсе — исправьте список или класс бетона перед переходом к клиенту.",
    },
  },
  bridge_piles: {
    productType: "bridge_piles",
    inputStep: "bridge_piles",
    endpointSegment: "bridge-piles",
    batchesField: "bridge_pile_batches",
    supportsGrades: true,
    isSimpleKp: true,
    ingestAction: "ingest_bridge_piles",
    labels: {
      nounPlural: "Мостовые сваи",
      nounGenitivePlural: "мостовых свай",
      stepTitle: "Шаг 1. Мостовые сваи",
      listLabel: "Список мостовых свай",
      reviewListTitle: "Список мостовых свай для расчёта",
      // Copy-paste leftover from the pile step — kept verbatim until the customer
      // confirms the canonical example (plan 2026-09-08 §5).
      placeholder: "С120.35-12 B25 5\nС120.35-13и 3",
      aiPlaceholder: "Например: убери строки с B15",
      emptySubtitle: "Вставьте текст списка мостовых свай или загрузите фото таблицы.",
      aiHint: "Редкий сценарий: опишите, что сделать со списком мостовых свай.",
      addMoreDescription: "Добавьте ещё мостовые сваи или перейдите к оформлению клиента.",
      initialDescription: "Загрузите фото или вставьте список мостовых свай для расчёта.",
      previewChangedMessage: "Изменён список мостовых свай — нажмите «Список верен» для пересчёта состава.",
      previewEmptyMessage: "Список пуст — распознайте мостовые сваи.",
      previewSubtitle: "Марка, класс бетона, количество и цена — как в документе.",
      previewUnpricedMessage:
        "Не все марки найдены в прайсе — исправьте список или класс бетона перед переходом к клиенту.",
    },
  },
  fbs: {
    productType: "fbs",
    inputStep: "fbs",
    endpointSegment: "fbs",
    batchesField: "fbs_batches",
    supportsGrades: true,
    isSimpleKp: true,
    ingestAction: "ingest_fbs",
    labels: {
      nounPlural: "ФБС",
      nounGenitivePlural: "ФБС",
      stepTitle: "Шаг 1. ФБС",
      listLabel: "Список ФБС",
      reviewListTitle: "Список ФБС для расчёта",
      // Same deliberate copy-paste leftover as bridge_piles (plan 2026-09-08 §5).
      placeholder: "С120.35-12 B25 5\nС120.35-13и 3",
      aiPlaceholder: "Например: убери строки с B15",
      emptySubtitle: "Вставьте текст списка ФБС или загрузите фото таблицы.",
      aiHint: "Редкий сценарий: опишите, что сделать со списком ФБС.",
      addMoreDescription: "Добавьте ещё ФБС или перейдите к оформлению клиента.",
      initialDescription: "Загрузите фото или вставьте список ФБС для расчёта.",
      previewChangedMessage: "Изменён список ФБС — нажмите «Список верен» для пересчёта состава.",
      previewEmptyMessage: "Список пуст — распознайте ФБС.",
      previewSubtitle: "Марка, класс бетона, количество и цена — как в документе.",
      previewUnpricedMessage:
        "Не все марки найдены в прайсе — исправьте список или класс бетона перед переходом к клиенту.",
    },
  },
};

/**
 * Unknown / missing product types resolve to plates — keeps legacy drafts working.
 * hasOwnProperty (not `in`): junk like "constructor" must not match the prototype chain.
 */
export const getProductTypeConfig = (type: ProductType | null | undefined): ProductTypeConfig => {
  if (type && Object.prototype.hasOwnProperty.call(PRODUCT_TYPE_CONFIG, type)) {
    return PRODUCT_TYPE_CONFIG[type];
  }
  return PRODUCT_TYPE_CONFIG.plates;
};

export const INGEST_REQUIRED_MESSAGE = "Сначала распознайте и получите хотя бы одну позицию в заказе.";

export type ResolveGateAction = Extract<
  WizardNextRequiredAction,
  "resolve_wide_plates" | "resolve_invalid_widths" | "resolve_unpriced_plates"
>;

/**
 * «Нельзя перейти дальше» для plates-гейтов по серверному next_required_action.
 * Бэкенд выставляет resolve_* только когда соответствующие metadata-списки плит
 * не пусты, т.е. де-факто для plates-циклов; для остальных изделий эти действия
 * недостижимы и здесь значатся только ради полноты lookup'а.
 */
export const RESOLVE_GATE_MESSAGES: Record<ResolveGateAction, string> = {
  resolve_wide_plates: "Сначала примите решение по позициям шире стандартной.",
  resolve_invalid_widths: "Нестандартная ширина: замените на заводской рез или исключите позицию.",
  resolve_unpriced_plates: "Сначала примите решение по позициям без цены в прайсе.",
};
