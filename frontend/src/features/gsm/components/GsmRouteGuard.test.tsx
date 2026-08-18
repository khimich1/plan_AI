import { cleanup, render, screen } from "@testing-library/react";
import { MemoryRouter, Route, Routes } from "react-router";
import { afterEach, describe, expect, it, vi } from "vitest";
import { RequireRole } from "@/features/auth/components/RequireRole";
import type { AuthUser } from "@/features/auth/model/types";

const mockUseAuth = vi.fn();

vi.mock("@/features/auth/model/AuthProvider", () => ({
  useAuth: () => mockUseAuth(),
}));

const GsmContent = () => <div>Раздел ГСМ</div>;
const NewHome = () => <div>Commercial home</div>;
const ProductionHome = () => <div>Production home</div>;
const LogisticsHome = () => <div>Logistics home</div>;

const buildUser = (role: string): AuthUser => ({
  id: 1,
  username: "test-user",
  role,
  manager_id: null,
  is_active: true,
});

const renderGuardedGsmRoute = (role: string) => {
  mockUseAuth.mockReturnValue({
    user: buildUser(role),
    isLoading: false,
    isAuthenticated: true,
  });

  return render(
    <MemoryRouter initialEntries={["/gsm"]}>
      <Routes>
        <Route path="/new" element={<NewHome />} />
        <Route path="/production" element={<ProductionHome />} />
        <Route path="/logistics" element={<LogisticsHome />} />
        <Route element={<RequireRole allowedRoles={["admin", "accountant"]} />}>
          <Route path="/gsm" element={<GsmContent />} />
        </Route>
      </Routes>
    </MemoryRouter>,
  );
};

describe("GSM route guard", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("renders GSM for admin", () => {
    renderGuardedGsmRoute("admin");
    expect(screen.getByText("Раздел ГСМ")).toBeInTheDocument();
  });

  it("renders GSM for accountant", () => {
    renderGuardedGsmRoute("accountant");
    expect(screen.getByText("Раздел ГСМ")).toBeInTheDocument();
  });

  it("redirects manager to commercial default route", () => {
    renderGuardedGsmRoute("manager");
    expect(screen.queryByText("Раздел ГСМ")).not.toBeInTheDocument();
    expect(screen.getByText("Commercial home")).toBeInTheDocument();
  });

  it("redirects production to production default route", () => {
    renderGuardedGsmRoute("production");
    expect(screen.queryByText("Раздел ГСМ")).not.toBeInTheDocument();
    expect(screen.getByText("Production home")).toBeInTheDocument();
  });

  it("redirects logistics to logistics default route", () => {
    renderGuardedGsmRoute("logistics");
    expect(screen.queryByText("Раздел ГСМ")).not.toBeInTheDocument();
    expect(screen.getByText("Logistics home")).toBeInTheDocument();
  });
});
