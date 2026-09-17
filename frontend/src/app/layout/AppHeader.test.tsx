import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { MemoryRouter } from "react-router";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { AppHeader } from "@/app/layout/AppHeader";

const mockUseAuth = vi.fn();
const mockUseCommercialDraftHeaderBridge = vi.fn();

vi.mock("@/features/auth/model/AuthProvider", () => ({
  useAuth: () => mockUseAuth(),
}));

vi.mock("@/pages/commercial-offer-create/CommercialOfferHeaderBridge", () => ({
  useCommercialDraftHeaderBridge: () => mockUseCommercialDraftHeaderBridge(),
}));

vi.mock("@/features/admin/components/DbManagementModal", () => ({
  DbManagementModal: ({
    open,
    onOpenImport1c,
  }: {
    open: boolean;
    onOpenImport1c?: () => void;
  }) =>
    open ? (
      <button type="button" onClick={onOpenImport1c}>
        Загрузить выгрузку 1С
      </button>
    ) : null,
}));

vi.mock("@/features/nomenclature/components/Import1cDialog", () => ({
  Import1cDialog: ({ open }: { open: boolean }) =>
    open ? <div data-testid="import-1c-dialog">import-1c-dialog</div> : null,
}));

const mockNavigate = vi.fn();
vi.mock("react-router", async () => {
  const actual = await vi.importActual<typeof import("react-router")>("react-router");
  return {
    ...actual,
    useNavigate: () => mockNavigate,
  };
});

vi.mock("@/features/notifications/components/NotificationBell", () => ({
  NotificationBell: () => <div data-testid="notification-bell-stub" />,
}));

describe("AppHeader nav", () => {
  beforeEach(() => {
    mockUseCommercialDraftHeaderBridge.mockReturnValue({
      hasDraft: false,
      resetDraft: vi.fn(),
    });
    mockUseAuth.mockReturnValue({
      user: { id: 1, role: "manager", username: "manager" },
      logout: vi.fn(),
      isLoggingOut: false,
    });
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("shows «Конструктор КП» nav label instead of «Создать КП»", () => {
    render(
      <MemoryRouter initialEntries={["/archive"]}>
        <AppHeader />
      </MemoryRouter>,
    );

    expect(screen.getByRole("link", { name: "Конструктор КП" })).toBeInTheDocument();
    expect(screen.queryByRole("link", { name: "Создать КП" })).not.toBeInTheDocument();
  });

  it("does not show 1C import in the header for manager", () => {
    render(
      <MemoryRouter initialEntries={["/archive"]}>
        <AppHeader />
      </MemoryRouter>,
    );

    expect(screen.queryByRole("button", { name: "Выгрузка 1С" })).not.toBeInTheDocument();
  });

  it("hides 1C import from production users", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 2, role: "production", username: "prod" },
      logout: vi.fn(),
      isLoggingOut: false,
    });
    render(
      <MemoryRouter initialEntries={["/production"]}>
        <AppHeader />
      </MemoryRouter>,
    );

    expect(screen.queryByRole("button", { name: "Выгрузка 1С" })).not.toBeInTheDocument();
  });

  it("hides 1C import and commercial nav from economist", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 6, role: "economist", username: "ekonomist" },
      logout: vi.fn(),
      isLoggingOut: false,
    });
    render(
      <MemoryRouter initialEntries={["/prices"]}>
        <AppHeader />
      </MemoryRouter>,
    );

    expect(screen.getByRole("link", { name: "Прайсы" })).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Выгрузка 1С" })).not.toBeInTheDocument();
    expect(screen.queryByRole("link", { name: "Конструктор КП" })).not.toBeInTheDocument();
    expect(screen.queryByRole("link", { name: "ГСМ" })).not.toBeInTheDocument();
  });

  it("shows prices nav for admin", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, role: "admin", username: "admin" },
      logout: vi.fn(),
      isLoggingOut: false,
    });
    render(
      <MemoryRouter initialEntries={["/archive"]}>
        <AppHeader />
      </MemoryRouter>,
    );

    expect(screen.getByRole("link", { name: "Прайсы" })).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Выгрузка 1С" })).not.toBeInTheDocument();
  });

  it("opens prices tab from admin DB settings", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, role: "admin", username: "admin" },
      logout: vi.fn(),
      isLoggingOut: false,
    });
    render(
      <MemoryRouter initialEntries={["/archive"]}>
        <AppHeader />
      </MemoryRouter>,
    );

    fireEvent.click(screen.getByRole("button", { name: "Управление БД" }));
    fireEvent.click(screen.getByRole("button", { name: "Загрузить выгрузку 1С" }));
    expect(mockNavigate).toHaveBeenCalledWith("/prices");
    expect(screen.queryByTestId("import-1c-dialog")).not.toBeInTheDocument();
  });
});
