import { cleanup, render, screen } from "@testing-library/react";
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
  DbManagementModal: () => null,
}));

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
});
