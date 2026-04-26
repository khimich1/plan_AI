import { useQuery } from "@tanstack/react-query";
import { authApi } from "@/features/auth/api/authApi";
import { ApiError } from "@/shared/lib/apiError";
import type { CurrentUser } from "@/features/auth/types/user";

export const authKeys = {
  me: ["auth", "me"] as const,
};

export const useCurrentUserQuery = () =>
  useQuery<CurrentUser | null>({
    queryKey: authKeys.me,
    queryFn: async () => {
      try {
        return await authApi.me();
      } catch (error) {
        if (error instanceof ApiError && error.status === 401) {
          return null;
        }
        throw error;
      }
    },
    staleTime: Infinity,
    retry: false,
  });
