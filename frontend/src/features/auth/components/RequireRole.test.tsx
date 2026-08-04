import { cleanup, render, screen } from "@testing-library/react";
import { MemoryRouter, Route, Routes } from "react-router";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { AuthUser } from "@/features/auth/model/types";
import { RequireRole } from "./RequireRole";

const mockUseAuth = vi.fn();

vi.mock("@/features/auth/model/AuthProvider", () => ({
  useAuth: () => mockUseAuth(),
}));

const ProtectedContent = () => <div>Protected content</div>;
const NewHome = () => <div>New home</div>;
const ProductionHome = () => <div>Production home</div>;

const buildUser = (role: string): AuthUser => ({
  id: 1,
  username: "test-user",
  role,
  manager_id: null,
  is_active: true,
});

const renderGuardedRoute = (role: string, initialPath = "/restricted") => {
  mockUseAuth.mockReturnValue({
    user: buildUser(role),
    isLoading: false,
    isAuthenticated: true,
  });

  return render(
    <MemoryRouter initialEntries={[initialPath]}>
      <Routes>
        <Route path="/new" element={<NewHome />} />
        <Route path="/production" element={<ProductionHome />} />
        <Route element={<RequireRole allowedRoles={["admin", "manager"]} />}>
          <Route path="/restricted" element={<ProtectedContent />} />
        </Route>
      </Routes>
    </MemoryRouter>,
  );
};

describe("RequireRole", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("renders child route when user role is allowed", () => {
    renderGuardedRoute("admin");
    expect(screen.getByText("Protected content")).toBeInTheDocument();
  });

  it("renders child route for manager on commercial guard", () => {
    renderGuardedRoute("manager");
    expect(screen.getByText("Protected content")).toBeInTheDocument();
  });

  it("redirects production users to their default route", () => {
    renderGuardedRoute("production");
    expect(screen.queryByText("Protected content")).not.toBeInTheDocument();
    expect(screen.getByText("Production home")).toBeInTheDocument();
  });

  it("redirects unknown roles to commercial default route", () => {
    renderGuardedRoute("guest");
    expect(screen.queryByText("Protected content")).not.toBeInTheDocument();
    expect(screen.getByText("New home")).toBeInTheDocument();
  });

  it("redirects when user is missing", () => {
    mockUseAuth.mockReturnValue({
      user: null,
      isLoading: false,
      isAuthenticated: false,
    });

    render(
      <MemoryRouter initialEntries={["/restricted"]}>
        <Routes>
          <Route path="/new" element={<NewHome />} />
          <Route element={<RequireRole allowedRoles={["admin", "manager"]} />}>
            <Route path="/restricted" element={<ProtectedContent />} />
          </Route>
        </Routes>
      </MemoryRouter>,
    );

    expect(screen.queryByText("Protected content")).not.toBeInTheDocument();
    expect(screen.getByText("New home")).toBeInTheDocument();
  });
});
