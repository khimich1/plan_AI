import { BrowserRouter, Navigate, Route, Routes } from "react-router-dom";
import { AppLayout } from "@/app/layout/AppLayout";
import { CommercialOfferCreatePage } from "@/pages/commercial-offer-create/CommercialOfferCreatePage";
import { CommercialOfferArchivePage } from "@/pages/commercial-offer-archive/CommercialOfferArchivePage";

export const AppRouter = () => (
  <BrowserRouter>
    <Routes>
      <Route element={<AppLayout />}>
        <Route path="/commercial-offer/new" element={<CommercialOfferCreatePage />} />
        <Route path="/commercial-offer/archive" element={<CommercialOfferArchivePage />} />
        <Route path="*" element={<Navigate to="/commercial-offer/new" replace />} />
      </Route>
    </Routes>
  </BrowserRouter>
);
