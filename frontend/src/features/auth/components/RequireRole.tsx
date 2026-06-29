import { Navigate, Outlet } from "react-router-dom";
import { useAuth } from "@/features/auth/model/AuthProvider";
import { defaultRouteForRole } from "@/shared/lib/roleRoutes";
import type { UserRole } from "@/features/auth/types/user";

type RequireRoleProps = {
  allowedRoles: readonly UserRole[];
};

export const RequireRole = ({ allowedRoles }: RequireRoleProps) => {
  const { user } = useAuth();
  const role = user?.role;

  if (!role || !allowedRoles.includes(role)) {
    return <Navigate to={defaultRouteForRole(role)} replace />;
  }

  return <Outlet />;
};
