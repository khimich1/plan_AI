import type { UserRole } from "@/features/auth/types/user";

/** UI hide ≠ authorization: role-based nav is UX only; API enforces RBAC server-side. */

const DEFAULT_COMMERCIAL_ROUTE = "/new";
const DEFAULT_PRODUCTION_ROUTE = "/production";
const DEFAULT_LOGISTICS_ROUTE = "/logistics";
const DEFAULT_GSM_ROUTE = "/gsm";

const COMMERCIAL_ROLES = ["admin", "manager"] as const satisfies readonly UserRole[];
const PRODUCTION_ROLES = ["admin", "production"] as const satisfies readonly UserRole[];
const LOGISTICS_ROLES = ["admin", "logistics"] as const satisfies readonly UserRole[];
const GSM_ROLES = ["admin", "accountant"] as const satisfies readonly UserRole[];

export const ROUTE_ACCESS: Record<string, readonly UserRole[]> = {
  "/new": COMMERCIAL_ROLES,
  "/archive": COMMERCIAL_ROLES,
  "/production": PRODUCTION_ROLES,
  "/logistics": LOGISTICS_ROLES,
  "/logistics/carriers": LOGISTICS_ROLES,
  "/gsm": GSM_ROLES,
};

function normalizeRoutePath(path: string): string {
  const withLeadingSlash = path.startsWith("/") ? path : `/${path}`;
  return withLeadingSlash.replace(/\/$/, "") || "/";
}

export function canAccessRoute(role: UserRole | undefined, path: string): boolean {
  const normalized = normalizeRoutePath(path);
  const allowedRoles = ROUTE_ACCESS[normalized];
  if (!allowedRoles) {
    return true;
  }
  if (!role) {
    return false;
  }
  return allowedRoles.includes(role);
}

export function defaultRouteForRole(role: UserRole | undefined): string {
  if (role === "production") {
    return DEFAULT_PRODUCTION_ROUTE;
  }
  if (role === "logistics") {
    return DEFAULT_LOGISTICS_ROUTE;
  }
  if (role === "accountant") {
    return DEFAULT_GSM_ROUTE;
  }
  return DEFAULT_COMMERCIAL_ROUTE;
}
