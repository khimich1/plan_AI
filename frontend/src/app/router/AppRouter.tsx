import { BrowserRouter, Navigate, Route, Routes } from "react-router-dom";
import { AppLayout } from "@/app/layout/AppLayout";
import { ProtectedRoute } from "@/features/auth/components/ProtectedRoute";
import { LoginPage } from "@/pages/login/LoginPage";
import { CommercialOfferCreatePage } from "@/pages/commercial-offer-create/CommercialOfferCreatePage";
import { CommercialOfferArchivePage } from "@/pages/commercial-offer-archive/CommercialOfferArchivePage";
import { ProductionPage } from "@/pages/production/ProductionPage";

export const AppRouter = () => (
  <BrowserRouter>
    <Routes>
      <Route path="/login" element={<LoginPage />} />
      <Route element={<ProtectedRoute />}>
        <Route element={<AppLayout />}>
          <Route path="/commercial-offer/new" element={<CommercialOfferCreatePage />} />
          <Route path="/commercial-offer/archive" element={<CommercialOfferArchivePage />} />
          <Route path="/production" element={<ProductionPage />} />
          <Route path="*" element={<Navigate to="/commercial-offer/new" replace />} />
        </Route>
      </Route>
    </Routes>
  </BrowserRouter>
);
