import { httpClient } from "@/shared/api/httpClient";
import type {
  PileCostApplyRequest,
  PileCostInputsBody,
  PileCostPreview,
  PileCostSaved,
  PileCostState,
} from "@/features/price-desk/types/pileCost";

const BASE = "/api/v1/prices/pile-cost";
const JSON_HEADERS = { "Content-Type": "application/json" };

const selectedMark = (mark?: string | null): string => mark?.trim() ?? "";

/** Только поля ручек. Объём и килограммы арматуры в JSON не попадают. */
const inputsPayload = (body: PileCostInputsBody): string =>
  JSON.stringify({
    prices: body.prices.map((row) => ({ article: row.article, price: row.price })),
    recipes: body.recipes.map((row) => ({
      concrete_grade: row.concrete_grade,
      cement_kg_m3: row.cement_kg_m3,
      sand_kg_m3: row.sand_kg_m3,
      crushed_kg_m3: row.crushed_kg_m3,
    })),
    overhead: {
      general: body.overhead.general,
      payroll: body.overhead.payroll,
      forming: body.overhead.forming,
      other: body.overhead.other,
    },
  });

export const pileCostApi = {
  get: (mark?: string | null) => {
    const markName = selectedMark(mark);
    const path = markName ? `${BASE}?mark=${encodeURIComponent(markName)}` : BASE;
    return httpClient.get<PileCostState>(path);
  },

  /** Сохраняет ручки. pile_prices не пишет. */
  save: (body: PileCostInputsBody) =>
    httpClient.put<PileCostSaved>(BASE, inputsPayload(body), JSON_HEADERS),

  /** Дифф текущих цен. Записи нет. */
  preview: () => httpClient.post<PileCostPreview>(`${BASE}/preview`),

  /** Запись в прайс только с отпечатком, который вернул превью. */
  apply: (body: PileCostApplyRequest) =>
    httpClient.post<PileCostPreview>(
      `${BASE}/apply`,
      JSON.stringify({ inputs_sha256: body.inputs_sha256 }),
      JSON_HEADERS,
    ),
};
