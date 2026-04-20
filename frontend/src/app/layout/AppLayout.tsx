import { Outlet } from "react-router-dom";
import { AppHeader } from "@/app/layout/AppHeader";

export const AppLayout = () => (
  <div className="app-shell">
    <AppHeader />
    <Outlet />
  </div>
);
