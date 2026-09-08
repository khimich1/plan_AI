import { httpClient } from "@/shared/api/httpClient";
import { ApiError } from "@/shared/lib/apiError";

const BASE = "/api/v1/counterparties";
const JSON_HEADERS = { "Content-Type": "application/json" };

export type CounterpartyShort = {
  id: number;
  code_1c: string;
  name: string;
  inn: string | null;
  kpp: string | null;
};

export type CounterpartySearchResponse = {
  items: CounterpartyShort[];
  count: number;
};

export type CounterpartyCreatePayload = {
  name: string;
  code_1c: string;
  inn?: string | null;
  kpp?: string | null;
  is_client?: boolean;
};

export type CounterpartyCreateResponse = {
  item: CounterpartyShort;
  warning?: string | null;
};

export class DuplicateCounterpartyError extends Error {
  existing: CounterpartyShort;

  constructor(message: string, existing: CounterpartyShort) {
    super(message);
    this.name = "DuplicateCounterpartyError";
    this.existing = existing;
  }
}

const asCounterpartyShort = (value: unknown): CounterpartyShort | null => {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }
  const record = value as Record<string, unknown>;
  if (typeof record.id !== "number" || typeof record.name !== "string" || typeof record.code_1c !== "string") {
    return null;
  }
  return {
    id: record.id,
    name: record.name,
    code_1c: record.code_1c,
    inn: typeof record.inn === "string" ? record.inn : null,
    kpp: typeof record.kpp === "string" ? record.kpp : null,
  };
};

const existingFromError = (error: unknown): CounterpartyShort | null => {
  if (!(error instanceof ApiError) || error.status !== 409) {
    return null;
  }
  const details = error.details;
  if (!details || typeof details !== "object" || Array.isArray(details)) {
    return null;
  }
  return asCounterpartyShort((details as Record<string, unknown>).existing);
};

export const searchCounterparties = (q: string): Promise<CounterpartySearchResponse> => {
  const params = new URLSearchParams({ q, limit: "10" });
  return httpClient.get<CounterpartySearchResponse>(`${BASE}/search?${params.toString()}`);
};

export const createCounterparty = async (
  payload: CounterpartyCreatePayload,
): Promise<CounterpartyCreateResponse> => {
  try {
    return await httpClient.post<CounterpartyCreateResponse>(
      `${BASE}`,
      JSON.stringify(payload),
      JSON_HEADERS,
    );
  } catch (error) {
    const existing = existingFromError(error);
    if (existing) {
      const message = error instanceof ApiError ? error.message : "Контрагент с таким кодом 1С уже есть";
      throw new DuplicateCounterpartyError(message, existing);
    }
    throw error;
  }
};
