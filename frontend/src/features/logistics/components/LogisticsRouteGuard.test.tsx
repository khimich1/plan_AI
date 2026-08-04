import { cleanup, render, screen } from "@testing-library/react";
import { MemoryRouter, Route, Routes } from "react-router";
import { afterEach, describe, expect, it, vi } from "vitest";
import { RequireRole } from "@/features/auth/components/RequireRole";
import type { AuthUser } from "@/features/auth/model/types";

const mockUseAuth = vi.fn();

vi.mock("@/features/auth/model/AuthProvider", () => ({
  useAuth: () => mockUseAuth(),
}));

const LogisticsContent = () => <div>Реестр рейсов</div>;
const NewHome = () => <div>Commercial home</div>;
const ProductionHome = () => <div>Production home</div>;

const buildUser = (role: string): AuthUser => ({
  id: 1,
  username: "test-user",
  role,
  manager_id: null,
  is_active: true,
});

const renderLogisticsRoute = (role: string) => {
  mockUseAuth.mockReturnValue({
    user: buildUser(role),
    isLoading: false,
    isAuthenticated: true,
  });

  return render(
    <MemoryRouter initialEntries={["/logistics"]}>
      <Routes>
        <Route path="/new" element={<NewHome />} />
        <Route path="/production" element={<ProductionHome />} />
        <Route element={<RequireRole allowedRoles={["admin", "logistics"]} />}>
          <Route path="/logistics" element={<LogisticsContent />} />
        </Route>
      </Routes>
    </MemoryRouter>,
  );
};

describe("Logistics route guard", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("renders logistics section for admin", () => {
    renderLogisticsRoute("admin");
    expect(screen.getByText("Реестр рейсов")).toBeInTheDocument();
  });

  it("renders logistics section for logistics role", () => {
    renderLogisticsRoute("logistics");
    expect(screen.getByText("Реестр рейсов")).toBeInTheDocument();
  });

  it("redirects manager to commercial default route", () => {
    renderLogisticsRoute("manager");
    expect(screen.queryByText("Реестр рейсов")).not.toBeInTheDocument();
    expect(screen.getByText("Commercial home")).toBeInTheDocument();
  });

  it("redirects production to production default route", () => {
    renderLogisticsRoute("production");
    expect(screen.queryByText("Реестр рейсов")).not.toBeInTheDocument();
    expect(screen.getByText("Production home")).toBeInTheDocument();
  });
});
