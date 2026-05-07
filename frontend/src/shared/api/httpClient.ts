import { env } from "@/shared/config/env";
import { ApiError } from "@/shared/lib/apiError";
import { queryClient } from "@/shared/lib/queryClient";

const AUTH_ME_QUERY_KEY = ["auth", "me"] as const;
const AUTH_LOGIN_PATH = "/api/v1/auth/login";

type RequestOptions = {
  method?: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
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

const looksLikeJson = (contentType: string, body: string): boolean => {
  const t = contentType.toLowerCase();
  if (t.includes("application/json") || t.includes("application/problem+json")) {
    return true;
  }
  const s = body.trimStart();
  return s.startsWith("{") || s.startsWith("[");
};

const parseResponseJson = async <T>(response: Response, path: string): Promise<T> => {
  const text = await response.text();
  const contentType = response.headers.get("Content-Type") ?? "";
  if (!looksLikeJson(contentType, text)) {
    const hint =
      text.trimStart().startsWith("<") || contentType.includes("text/html")
        ? " Похоже на HTML (часто SPA Vite без proxy на /api): проверьте server.proxy и что бэкенд слушает порт из Vite."
        : "";
    throw new ApiError(
      `Ожидался JSON для ${path}, получен «${contentType || "нет Content-Type"}».${hint}`,
      response.status,
    );
  }
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new ApiError(`Не удалось разобрать JSON для ${path}.`, response.status);
  }
};

const parseError = async (response: Response, path: string): Promise<never> => {
  let detail = "Запрос завершился ошибкой.";
  const contentType = response.headers.get("Content-Type") ?? "";
  try {
    const text = await response.text();
    if (looksLikeJson(contentType, text)) {
      try {
        const payload = JSON.parse(text) as { detail?: string };
        if (typeof payload.detail === "string" && payload.detail.trim()) {
          detail = payload.detail;
        }
      } catch {
        detail = response.statusText || detail;
      }
    } else if (text.trimStart().startsWith("<") || contentType.includes("text/html")) {
      detail = `Сервер вернул HTML вместо JSON (${response.status}) для ${path}. Проверьте URL API и прокси dev-сервера.`;
    }
  } catch {
    detail = response.statusText || detail;
  }
  throw new ApiError(detail, response.status, detail);
};

const handleUnauthorized = (path: string): void => {
  if (path.startsWith(AUTH_LOGIN_PATH)) {
    return;
  }
  queryClient.setQueryData(AUTH_ME_QUERY_KEY, null);
  queryClient.invalidateQueries({ queryKey: AUTH_ME_QUERY_KEY });
};

const request = async <TResponse>(path: string, options: RequestOptions = {}): Promise<TResponse> => {
  const response = await fetch(buildUrl(path), {
    method: options.method ?? "GET",
    body: options.body ?? null,
    headers: options.headers,
    credentials: "include",
  });

  if (!response.ok) {
    if (response.status === 401) {
      handleUnauthorized(path);
    }
    return parseError(response, path);
  }

  if (response.status === 204) {
    return undefined as TResponse;
  }

  return parseResponseJson<TResponse>(response, path);
};

export type DownloadResult = {
  blob: Blob;
  filename: string;
  contentType: string;
};

const extractFilename = (response: Response, fallback: string): string => {
  const header = response.headers.get("Content-Disposition");
  if (!header) {
    return fallback;
  }
  const utf8Match = header.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match) {
    try {
      return decodeURIComponent(utf8Match[1]);
    } catch {
      /* fall through */
    }
  }
  const plainMatch = header.match(/filename="?([^";]+)"?/i);
  if (plainMatch) {
    return plainMatch[1];
  }
  return fallback;
};

const downloadRequest = async (
  path: string,
  fallbackFilename: string,
): Promise<DownloadResult> => {
  const response = await fetch(buildUrl(path), {
    method: "GET",
    credentials: "include",
  });
  if (!response.ok) {
    if (response.status === 401) {
      handleUnauthorized(path);
    }
    return parseError(response, path);
  }
  const blob = await response.blob();
  const filename = extractFilename(response, fallbackFilename);
  const contentType = response.headers.get("Content-Type") ?? blob.type;
  return { blob, filename, contentType };
};

export const httpClient = {
  get: <TResponse>(path: string) => request<TResponse>(path),
  post: <TResponse>(path: string, body?: BodyInit | null, headers?: HeadersInit) =>
    request<TResponse>(path, { method: "POST", body, headers }),
  put: <TResponse>(path: string, body?: BodyInit | null, headers?: HeadersInit) =>
    request<TResponse>(path, { method: "PUT", body, headers }),
  patch: <TResponse>(path: string, body?: BodyInit | null, headers?: HeadersInit) =>
    request<TResponse>(path, { method: "PATCH", body, headers }),
  delete: <TResponse>(path: string) => request<TResponse>(path, { method: "DELETE" }),
  request: <TResponse>(path: string, options: RequestOptions) =>
    request<TResponse>(path, options),
  download: (path: string, fallbackFilename = "download"): Promise<DownloadResult> =>
    downloadRequest(path, fallbackFilename),
};

export const resolveApiUrl = (path: string): string => buildUrl(path);
