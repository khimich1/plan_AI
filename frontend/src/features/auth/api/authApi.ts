import { httpClient } from "@/shared/api/httpClient";
import { ApiError } from "@/shared/lib/apiError";
import type {
  AuthUser,
  LoginPayload,
  LoginResponse,
  MeResponse,
} from "@/features/auth/model/types";

const BASE = "/api/v1/auth";

export const authApi = {
  login: (payload: LoginPayload) =>
    httpClient.post<LoginResponse>(
      `${BASE}/login`,
      JSON.stringify(payload),
      { "Content-Type": "application/json" },
    ),

  logout: () => httpClient.post<{ ok: boolean }>(`${BASE}/logout`),

  me: async (): Promise<AuthUser | null> => {
    try {
      const data = await httpClient.get<MeResponse>(`${BASE}/me`);
      return data.user;
    } catch (error) {
      if (error instanceof ApiError && error.status === 401) {
        return null;
      }
      throw error;
    }
  },
};
