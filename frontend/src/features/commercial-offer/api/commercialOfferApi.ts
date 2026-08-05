import { httpClient } from "@/shared/api/httpClient";
import type {
  BreakdownResponse,
  CommercialDraftDetails,
  CommercialGeneratedFile,
  CommercialSaveResult,
  ConditionsMode,
  ManagersResponse,
  PlateInputMode,
  ProductType,
  SaveMode,
  WidePlateAction,
} from "@/features/commercial-offer/types/commercialOffer";

type DraftCreatePayload = {
  text: string;
  image: File | null;
  productType?: ProductType;
};

type UpdateDraftPlatesPayload = DraftCreatePayload & {
  mode: PlateInputMode;
};

type UpdateDraftPilesPayload = UpdateDraftPlatesPayload;

type UpdateDraftStepsPayload = UpdateDraftPlatesPayload;

type UpdateDraftMarchesPayload = UpdateDraftPlatesPayload;

type UpdateDraftMetaPayload = {
  managerId?: number | null;
  clientName?: string;
  discountPercent?: number;
  conditionsMode?: ConditionsMode;
  deliveryConditions?: string;
  paymentConditions?: string;
  logisticsCost?: number;
};

type WidePlateDecisionPayload = {
  lineId?: string;
  sourceLine: string;
  action: WidePlateAction;
  replacementText?: string;
};

type SaveDraftPayload = {
  mode: SaveMode;
  executionTermsInput?: string;
};

const createMultipartPayload = ({
  text,
  image,
  mode,
  productType,
}: DraftCreatePayload & { mode?: PlateInputMode }) => {
  const formData = new FormData();
  if (text.trim()) {
    formData.append("text", text.trim());
  }
  if (image) {
    formData.append("image", image);
  }
  if (mode) {
    formData.append("mode", mode);
  }
  if (productType) {
    formData.append("product_type", productType);
  }
  return formData;
};

type ApplyAiPlatesPayload = {
  instruction: string;
  image: File | null;
};

const createAiMultipartPayload = ({ instruction, image }: ApplyAiPlatesPayload) => {
  const formData = new FormData();
  formData.append("instruction", instruction.trim());
  if (image) {
    formData.append("image", image);
  }
  return formData;
};

export const commercialOfferApi = {
  getManagers: () => httpClient.get<ManagersResponse>("/api/v1/managers"),

  createDraft: (payload: DraftCreatePayload) =>
    httpClient.post<CommercialDraftDetails>("/api/v1/commercial/drafts", createMultipartPayload(payload)),

  updateDraftPlates: (draftId: string, payload: UpdateDraftPlatesPayload) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/plates`,
      createMultipartPayload(payload),
    ),

  updateDraftPiles: (draftId: string, payload: UpdateDraftPilesPayload) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/piles`,
      createMultipartPayload(payload),
    ),

  updateDraftSteps: (draftId: string, payload: UpdateDraftStepsPayload) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/steps`,
      createMultipartPayload(payload),
    ),

  updateDraftMarches: (draftId: string, payload: UpdateDraftMarchesPayload) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/marches`,
      createMultipartPayload(payload),
    ),

  updateDraftBridgePiles: (draftId: string, payload: UpdateDraftPilesPayload) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/bridge-piles`,
      createMultipartPayload(payload),
    ),

  updateDraftFbs: (draftId: string, payload: UpdateDraftPilesPayload) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/fbs`,
      createMultipartPayload(payload),
    ),

  applyAiPlates: (draftId: string, payload: ApplyAiPlatesPayload) =>
    httpClient.post<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/plates/ai`,
      createAiMultipartPayload(payload),
    ),

  applyAiPiles: (draftId: string, payload: ApplyAiPlatesPayload) =>
    httpClient.post<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/piles/ai`,
      createAiMultipartPayload(payload),
    ),

  applyAiSteps: (draftId: string, payload: ApplyAiPlatesPayload) =>
    httpClient.post<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/steps/ai`,
      createAiMultipartPayload(payload),
    ),

  applyAiMarches: (draftId: string, payload: ApplyAiPlatesPayload) =>
    httpClient.post<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/marches/ai`,
      createAiMultipartPayload(payload),
    ),

  applyAiBridgePiles: (draftId: string, payload: ApplyAiPlatesPayload) =>
    httpClient.post<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/bridge-piles/ai`,
      createAiMultipartPayload(payload),
    ),

  applyAiFbs: (draftId: string, payload: ApplyAiPlatesPayload) =>
    httpClient.post<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/fbs/ai`,
      createAiMultipartPayload(payload),
    ),

  updatePileGrades: (draftId: string, concreteGrade: string) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/piles/grades`,
      JSON.stringify({ concrete_grade: concreteGrade }),
      { "Content-Type": "application/json" },
    ),

  updateMarchGrades: (draftId: string, concreteGrade: string) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/marches/grades`,
      JSON.stringify({ concrete_grade: concreteGrade }),
      { "Content-Type": "application/json" },
    ),

  updateBridgePileGrades: (draftId: string, concreteGrade: string) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/bridge-piles/grades`,
      JSON.stringify({ concrete_grade: concreteGrade }),
      { "Content-Type": "application/json" },
    ),

  updateFbsGrades: (draftId: string, concreteGrade: string) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/fbs/grades`,
      JSON.stringify({ concrete_grade: concreteGrade }),
      { "Content-Type": "application/json" },
    ),

  resolveWidePlates: (draftId: string, decisions: WidePlateDecisionPayload[]) =>
    httpClient.post<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/wide-plates/resolve`,
      JSON.stringify({
        decisions: decisions.map((item) => ({
          line_id: item.lineId ?? null,
          source_line: item.sourceLine,
          action: item.action,
          replacement_text: item.replacementText ?? "",
        })),
      }),
      { "Content-Type": "application/json" },
    ),

  updateDraftMeta: (draftId: string, payload: UpdateDraftMetaPayload) =>
    httpClient.patch<CommercialDraftDetails>(
      `/api/v1/commercial/drafts/${draftId}/meta`,
      JSON.stringify({
        manager_id: payload.managerId,
        client_name: payload.clientName,
        discount_percent: payload.discountPercent,
        conditions_mode: payload.conditionsMode,
        delivery_conditions: payload.deliveryConditions,
        payment_conditions: payload.paymentConditions,
        logistics_cost: payload.logisticsCost,
      }),
      { "Content-Type": "application/json" },
    ),

  getDraft: (draftId: string) => httpClient.get<CommercialDraftDetails>(`/api/v1/commercial/drafts/${draftId}`),

  getBreakdown: (draftId: string) =>
    httpClient.get<BreakdownResponse>(`/api/v1/commercial/drafts/${draftId}/breakdown`),

  calculateDraft: (draftId: string) =>
    httpClient.post<CommercialDraftDetails>(`/api/v1/commercial/drafts/${draftId}/calculate`),

  generateFiles: (draftId: string, fileTypes?: CommercialGeneratedFile["kind"][]) =>
    httpClient.post<{ draft_id: string; files: CommercialGeneratedFile[] }>(
      `/api/v1/commercial/drafts/${draftId}/generate-files`,
      JSON.stringify({ file_types: fileTypes ?? ["pdf", "xlsx", "breakdown"] }),
      { "Content-Type": "application/json" },
    ),

  generateSchemaFiles: (draftId: string) =>
    httpClient.post<{ draft_id: string; files: CommercialGeneratedFile[] }>(
      `/api/v1/commercial/drafts/${draftId}/generate-files`,
      JSON.stringify({ file_types: ["schema"] }),
      { "Content-Type": "application/json" },
    ),

  saveDraft: (draftId: string, payload: SaveDraftPayload) =>
    httpClient.post<CommercialSaveResult>(
      `/api/v1/commercial/drafts/${draftId}/save`,
      JSON.stringify({
        mode: payload.mode,
        execution_terms_input: payload.executionTermsInput ?? "",
      }),
      { "Content-Type": "application/json" },
    ),
};
