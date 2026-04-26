import { Navigate, Outlet } from "react-router-dom";
import { useAuth } from "@/features/auth/model/AuthProvider";
import { Spinner } from "@/shared/ui/Spinner";

export const ProtectedRoute = () => {
  const { user, isLoading } = useAuth();

  if (isLoading) {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "grid",
          placeItems: "center",
        }}
      >
        <Spinner />
      </div>
    );
  }

  if (!user) {
    return <Navigate to="/login" replace />;
  }

  return <Outlet />;
};
