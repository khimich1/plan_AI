import { BrowserRouter, Navigate, Route, Routes } from "react-router-dom";
import { AppLayout } from "@/app/layout/AppLayout";
import { ProtectedRoute } from "@/features/auth/components/ProtectedRoute";
import { useAuth } from "@/features/auth/model/AuthProvider";
import { LoginPage } from "@/pages/login/LoginPage";
import { CommercialOfferCreatePage } from "@/pages/commercial-offer-create/CommercialOfferCreatePage";
import { CommercialOfferArchivePage } from "@/pages/commercial-offer-archive/CommercialOfferArchivePage";
import { ProductionPage } from "@/pages/production/ProductionPage";
import { defaultRouteForRole } from "@/shared/lib/roleRoutes";

const routerBasename = import.meta.env.BASE_URL.replace(/\/$/, "");

const RoleHomeRedirect = () => {
  const { user } = useAuth();
  return <Navigate to={defaultRouteForRole(user?.role)} replace />;
};

export const AppRouter = () => (
  <BrowserRouter basename={routerBasename}>
    <Routes>
      <Route path="/login" element={<LoginPage />} />
      <Route element={<ProtectedRoute />}>
        <Route element={<AppLayout />}>
          {/* Относительные сегменты: в RR7 надёжнее вложенности под pathless layout + basename */}
          <Route index element={<RoleHomeRedirect />} />
          <Route path="new" element={<CommercialOfferCreatePage />} />
          <Route path="archive" element={<CommercialOfferArchivePage />} />
          <Route path="production" element={<ProductionPage />} />
          <Route path="*" element={<RoleHomeRedirect />} />
        </Route>
      </Route>
    </Routes>
  </BrowserRouter>
);
