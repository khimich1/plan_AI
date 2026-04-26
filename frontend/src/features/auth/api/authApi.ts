import { httpClient } from "@/shared/api/httpClient";
import type { CurrentUserResponse } from "@/features/auth/types/user";

const BASE = "/api/v1/auth";

export const authApi = {
  me: () => httpClient.get<CurrentUserResponse>(`${BASE}/me`),
};
