import { env } from "@/shared/config/env";
import { ApiError } from "@/shared/lib/apiError";

type RequestOptions = {
  method?: "GET" | "POST" | "PATCH" | "DELETE";
  body?: BodyInit | null;
  headers?: HeadersInit;
};

const buildUrl = (path: string): string => {
  if (path.startsWith("http")) {
    return path;
  }
  const normalizedPath = path.startsWith("/") ? path : `/${path}`;
  return `${env.apiBaseUrl}${normalizedPath}`;
};

const parseError = async (response: Response): Promise<never> => {
  let detail = "Запрос завершился ошибкой.";
  try {
    const payload = (await response.json()) as { detail?: string };
    if (typeof payload.detail === "string" && payload.detail.trim()) {
      detail = payload.detail;
    }
  } catch {
    detail = response.statusText || detail;
  }
  throw new ApiError(detail, response.status, detail);
};

const request = async <TResponse>(path: string, options: RequestOptions = {}): Promise<TResponse> => {
  const response = await fetch(buildUrl(path), {
    method: options.method ?? "GET",
    body: options.body ?? null,
    headers: options.headers,
    credentials: "include",
  });

  if (!response.ok) {
    return parseError(response);
  }

  if (response.status === 204) {
    return undefined as TResponse;
  }

  return (await response.json()) as TResponse;
};

export const httpClient = {
  get: <TResponse>(path: string) => request<TResponse>(path),
  post: <TResponse>(path: string, body?: BodyInit | null, headers?: HeadersInit) =>
    request<TResponse>(path, { method: "POST", body, headers }),
  patch: <TResponse>(path: string, body?: BodyInit | null, headers?: HeadersInit) =>
    request<TResponse>(path, { method: "PATCH", body, headers }),
  delete: <TResponse>(path: string) => request<TResponse>(path, { method: "DELETE" }),
};

export const resolveApiUrl = (path: string): string => buildUrl(path);
