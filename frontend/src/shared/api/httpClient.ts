import { env } from "@/shared/config/env";
import { ApiError, parseApiErrorPayload } from "@/shared/lib/apiError";
import { queryClient } from "@/shared/lib/queryClient";

const AUTH_ME_QUERY_KEY = ["auth", "me"] as const;
const AUTH_LOGIN_PATH = "/api/v1/auth/login";
const AUTH_ME_PATH = "/api/v1/auth/me";
const CSRF_COOKIE_NAME = "csrf_token";
const CSRF_HEADER_NAME = "X-CSRF-Token";
const SAFE_METHODS = new Set(["GET", "HEAD", "OPTIONS", "TRACE"]);

const readCsrfToken = (): string | null => {
  if (typeof document === "undefined") {
    return null;
  }
  const prefix = `${CSRF_COOKIE_NAME}=`;
  const cookies = document.cookie.split(";");
  for (const raw of cookies) {
    const part = raw.trim();
    if (part.startsWith(prefix)) {
      return decodeURIComponent(part.slice(prefix.length));
    }
  }
  return null;
};

type RequestOptions = {
  method?: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  body?: BodyInit | null;
  headers?: HeadersInit;
  signal?: AbortSignal;
};

const buildUrl = (path: string): string => {
  if (path.startsWith("http")) {
    return path;
  }
  const normalizedPath = path.startsWith("/") ? path : `/${path}`;
  return `${env.apiBaseUrl}${normalizedPath}`;
};

/** Bootstrap CSRF cookie when missing (e.g. /auth/me failed before backend was up). */
const ensureCsrfToken = async (): Promise<string | null> => {
  const existing = readCsrfToken();
  if (existing) {
    return existing;
  }
  try {
    await fetch(buildUrl("/api/v1/health"), {
      method: "GET",
      credentials: "include",
    });
  } catch {
    return null;
  }
  return readCsrfToken();
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
  let message = "Запрос завершился ошибкой.";
  let code: string | undefined;
  let details: unknown;
  const contentType = response.headers.get("Content-Type") ?? "";
  try {
    const text = await response.text();
    if (looksLikeJson(contentType, text)) {
      try {
        const payload = JSON.parse(text) as unknown;
        const parsed = parseApiErrorPayload(payload);
        if (parsed) {
          message = parsed.message;
          code = parsed.code;
          details = parsed.details;
        }
      } catch {
        message = response.statusText || message;
      }
    } else if (text.trimStart().startsWith("<") || contentType.includes("text/html")) {
      message = `Сервер вернул HTML вместо JSON (${response.status}) для ${path}. Проверьте URL API и прокси dev-сервера.`;
    }
  } catch {
    message = response.statusText || message;
  }
  throw new ApiError(message, response.status, message, code, details);
};

const handleUnauthorized = (path: string): void => {
  // /auth/me itself returns 401 when logged out — do not invalidate or we loop:
  // me → 401 → invalidate → me → 401 …
  if (path.startsWith(AUTH_LOGIN_PATH) || path.startsWith(AUTH_ME_PATH)) {
    return;
  }
  queryClient.setQueryData(AUTH_ME_QUERY_KEY, null);
  queryClient.invalidateQueries({ queryKey: AUTH_ME_QUERY_KEY });
};

const request = async <TResponse>(path: string, options: RequestOptions = {}): Promise<TResponse> => {
  const method = options.method ?? "GET";
  const headers = new Headers(options.headers);
  if (!SAFE_METHODS.has(method)) {
    const csrfToken = await ensureCsrfToken();
    if (csrfToken) {
      headers.set(CSRF_HEADER_NAME, csrfToken);
    }
  }

  let response: Response;
  try {
    response = await fetch(buildUrl(path), {
      method,
      body: options.body ?? null,
      headers,
      credentials: "include",
      signal: options.signal,
    });
  } catch (err) {
    if (
      (err instanceof DOMException && err.name === "AbortError") ||
      (err instanceof Error && err.name === "AbortError") ||
      options.signal?.aborted
    ) {
      throw err;
    }
    throw new ApiError(
      "Сервер недоступен. Подождите пару секунд после запуска и обновите страницу.",
      0,
    );
  }

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

type DownloadRequestOptions = {
  method?: "GET" | "POST";
  body?: BodyInit | null;
  headers?: HeadersInit;
};

const downloadRequest = async (
  path: string,
  fallbackFilename: string,
  options: DownloadRequestOptions = {},
): Promise<DownloadResult> => {
  const method = options.method ?? "GET";
  const headers = new Headers(options.headers);
  if (!SAFE_METHODS.has(method)) {
    const csrfToken = await ensureCsrfToken();
    if (csrfToken) {
      headers.set(CSRF_HEADER_NAME, csrfToken);
    }
  }

  let response: Response;
  try {
    response = await fetch(buildUrl(path), {
      method,
      body: options.body ?? null,
      headers,
      credentials: "include",
    });
  } catch {
    throw new ApiError(
      "Сервер недоступен. Подождите пару секунд после запуска и обновите страницу.",
      0,
    );
  }

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
  post: <TResponse>(
    path: string,
    body?: BodyInit | null,
    headers?: HeadersInit,
    extra?: { signal?: AbortSignal },
  ) => request<TResponse>(path, { method: "POST", body, headers, signal: extra?.signal }),
  put: <TResponse>(path: string, body?: BodyInit | null, headers?: HeadersInit) =>
    request<TResponse>(path, { method: "PUT", body, headers }),
  patch: <TResponse>(path: string, body?: BodyInit | null, headers?: HeadersInit) =>
    request<TResponse>(path, { method: "PATCH", body, headers }),
  delete: <TResponse>(path: string) => request<TResponse>(path, { method: "DELETE" }),
  request: <TResponse>(path: string, options: RequestOptions) =>
    request<TResponse>(path, options),
  download: (
    path: string,
    fallbackFilename = "download",
    options?: DownloadRequestOptions,
  ): Promise<DownloadResult> => downloadRequest(path, fallbackFilename, options),
};

export const resolveApiUrl = (path: string): string => buildUrl(path);
