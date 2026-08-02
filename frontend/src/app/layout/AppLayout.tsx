import { Outlet } from "react-router";
import { AppHeader } from "@/app/layout/AppHeader";
import { CommercialOfferHeaderBridgeProvider } from "@/pages/commercial-offer-create/CommercialOfferHeaderBridge";

export const AppLayout = () => (
  <div className="app-shell">
    <CommercialOfferHeaderBridgeProvider>
      <AppHeader />
      <Outlet />
    </CommercialOfferHeaderBridgeProvider>
  </div>
);
