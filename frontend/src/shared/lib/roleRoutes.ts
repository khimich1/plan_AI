import type { UserRole } from "@/features/auth/types/user";

const DEFAULT_COMMERCIAL_ROUTE = "/new";
const DEFAULT_PRODUCTION_ROUTE = "/production";

export function defaultRouteForRole(role: UserRole | undefined): string {
  if (role === "production") {
    return DEFAULT_PRODUCTION_ROUTE;
  }
  return DEFAULT_COMMERCIAL_ROUTE;
}
