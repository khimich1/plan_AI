export type WizardStepId = "plates" | "client" | "result";

/** Legacy step ids from older drafts (localStorage / server metadata). */
export type LegacyWizardStepId = "wide-plates" | "manager";

/** Синхронизировано с app.schemas.commercial.WizardNextRequiredAction */
export type WizardNextRequiredAction =
  | "none"
  | "ingest_plates"
  | "resolve_wide_plates"
  | "select_manager"
  | "complete_client_terms"
  | "post_calculate";

/** Серверный контракт оркестрации (`CommercialWizardState` / OpenAPI `wizard_state`). */
export type CommercialWizardState = {
  current_step: WizardStepId;
  can_proceed_to: WizardStepId[];
  next_required_action: WizardNextRequiredAction;
  validation_errors: string[];
};

export type CommercialWizardStateResponse = CommercialWizardState;

export type PlateInputMode = "append" | "replace";
export type ConditionsMode = "standard" | "custom";
export type SaveMode = "database" | "archive" | "skip";
export type FileKind = "pdf" | "xlsx" | "breakdown" | "schema";
export type WidePlateAction = "confirm" | "exclude" | "replace";

export type Manager = {
  id: number;
  fio: string;
  contact_number: string;
  email: string;
};

export type ManagersResponse = {
  items: Manager[];
  count: number;
};

export type WidePlateLine = {
  id: string;
  line: string;
  qty: number;
};

export type DoborPair = {
  id: string;
  source_line: string;
  primary_line: string;
  complement_line: string;
};

export type PlateBatch = {
  source_type: "text" | "image" | "ai";
  original_text: string;
  normalized_text: string;
  ocr_text: string;
  filename: string;
};

export type CommercialGeneratedFile = {
  kind: FileKind;
  filename: string;
  display_name: string;
  download_url: string;
};

export type CommercialOfferIdentity = {
  offer_number: string;
  offer_date: string;
  file_stem: string;
};

export type CommercialSavedOffer = {
  kp_id: number | null;
  status: string;
  mode: SaveMode;
  execution_terms: string;
  saved_at: string;
};

export type OcrCorrection = {
  action: string;
  row_index?: number | null;
  before?: { normalized_candidate?: string; qty?: number } | null;
  after?: { normalized_candidate?: string; qty?: number } | null;
  reason?: string;
};

export type CommercialDraftMetadata = {
  source_type: "text" | "image" | "ai" | null;
  original_text: string;
  ocr_text: string;
  input_text: string;
  accumulated_text: string;
  manager_id: number | null;
  manager_name: string;
  manager_phone: string;
  manager_email: string;
  client_name: string;
  discount_percent: number;
  conditions_mode: ConditionsMode;
  delivery_conditions: string;
  payment_conditions: string;
  warnings: string[];
  unparsed_lines: string[];
  normalized_text: string;
  normalized_lines: string[];
  wide_plate_lines: WidePlateLine[];
  dobor_pairs: DoborPair[];
  diagnostics: Array<Record<string, unknown>>;
  price_rows_count: number;
  breakdown_tables_count: number;
  total_sum: number;
  plate_batches: PlateBatch[];
  wide_plates_resolved: boolean;
  last_source_filename: string;
  ai_applied?: boolean;
  last_ai_instruction?: string;
  current_step: string;
  current_save_mode: SaveMode | null;
  execution_terms: string;
  logistics_cost: number;
  ocr_method?: string;
  ocr_verify_applied?: boolean;
  ocr_verify_failed?: boolean;
  ocr_corrections?: OcrCorrection[];
  ocr_row_count_on_image?: number | null;
};

export type BreakdownTable = {
  name: string;
  rows: string[][];
};

export type BreakdownResponse = {
  draft_id: string;
  items: BreakdownTable[];
};

export type CommercialDraftDetails = {
  draft_id: string;
  order: Record<string, unknown>;
  optimization: {
    result?: Record<string, unknown>;
    total_plates: number;
    total_cost: number;
  };
  order_data: Array<Record<string, unknown>>;
  metadata: CommercialDraftMetadata;
  wizard_state: CommercialWizardState;
  files: CommercialGeneratedFile[];
  saved_offer: CommercialSavedOffer | null;
  totals: {
    total_qty?: number;
    subtotal?: number;
    vat_amount?: number;
    total_with_vat?: number;
  };
  offer_identity: CommercialOfferIdentity;
};

export type CommercialSaveResult = {
  draft_id: string;
  saved_offer: CommercialSavedOffer | null;
  totals: CommercialDraftDetails["totals"];
  offer_identity: CommercialOfferIdentity;
  result_card: {
    kp_id: number | null;
    offer_number: string;
    offer_date: string;
    client_name: string;
    manager_name: string;
    total_amount: number;
    status: string;
    execution_terms: string;
  };
};

export type WizardStoreState = {
  draftId: string | null;
  currentStep: WizardStepId;
  sourceText: string;
  selectedImageName: string | null;
  normalizedText: string;
  batchReviewText: string;
  pendingBatchReview: boolean;
  confirmedBatchCount: number;
  lastPlateMode: PlateInputMode;
  managerId: number | null;
  clientName: string;
  discountPercent: number;
  conditionsMode: ConditionsMode;
  deliveryConditions: string;
  paymentConditions: string;
  executionTermsInput: string;
  widePlateActions: Record<string, { action: WidePlateAction; replacementText: string }>;
  lastDraft: CommercialDraftDetails | null;
  lastSaveResult: CommercialSaveResult | null;
};
