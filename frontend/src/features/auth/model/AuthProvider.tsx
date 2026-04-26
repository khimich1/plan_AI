import { createContext, useContext, useMemo, type PropsWithChildren } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import { authApi } from "@/features/auth/api/authApi";
import type { AuthUser, LoginPayload } from "@/features/auth/model/types";

export const AUTH_ME_QUERY_KEY = ["auth", "me"] as const;

type AuthContextValue = {
  user: AuthUser | null;
  isLoading: boolean;
  isAuthenticated: boolean;
  login: (payload: LoginPayload) => Promise<AuthUser>;
  logout: () => Promise<void>;
  isLoggingIn: boolean;
  isLoggingOut: boolean;
};

const AuthContext = createContext<AuthContextValue | null>(null);

export const AuthProvider = ({ children }: PropsWithChildren) => {
  const queryClient = useQueryClient();

  const meQuery = useQuery({
    queryKey: AUTH_ME_QUERY_KEY,
    queryFn: authApi.me,
    staleTime: 5 * 60 * 1000,
    retry: false,
  });

  const loginMutation = useMutation({
    mutationFn: (payload: LoginPayload) => authApi.login(payload),
    onSuccess: (data) => {
      queryClient.setQueryData<AuthUser | null>(AUTH_ME_QUERY_KEY, data.user);
    },
  });

  const logoutMutation = useMutation({
    mutationFn: () => authApi.logout(),
    onSuccess: () => {
      queryClient.setQueryData<AuthUser | null>(AUTH_ME_QUERY_KEY, null);
      queryClient.clear();
    },
  });

  const value = useMemo<AuthContextValue>(
    () => ({
      user: meQuery.data ?? null,
      isLoading: meQuery.isLoading,
      isAuthenticated: Boolean(meQuery.data),
      isLoggingIn: loginMutation.isPending,
      isLoggingOut: logoutMutation.isPending,
      login: async (payload) => {
        const result = await loginMutation.mutateAsync(payload);
        return result.user;
      },
      logout: async () => {
        await logoutMutation.mutateAsync();
      },
    }),
    [
      meQuery.data,
      meQuery.isLoading,
      loginMutation,
      logoutMutation,
    ],
  );

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
};

export const useAuth = (): AuthContextValue => {
  const ctx = useContext(AuthContext);
  if (!ctx) {
    throw new Error("useAuth must be used within <AuthProvider>");
  }
  return ctx;
};
