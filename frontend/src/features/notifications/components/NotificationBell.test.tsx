import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { MemoryRouter } from "react-router";
import { NotificationBell } from "@/features/notifications/components/NotificationBell";
import type { NotificationItem, NotificationList } from "@/features/notifications/api/notifications";

const mockNavigate = vi.fn();
const mockMutate = vi.fn();

const listState: { data: NotificationList | undefined } = {
  data: { items: [], unread_count: 0 },
};

vi.mock("react-router", async () => {
  const actual = await vi.importActual<typeof import("react-router")>("react-router");
  return {
    ...actual,
    useNavigate: () => mockNavigate,
  };
});

vi.mock("@/features/notifications/api/notifications", async () => {
  const actual = await vi.importActual<typeof import("@/features/notifications/api/notifications")>(
    "@/features/notifications/api/notifications",
  );
  return {
    ...actual,
    useNotificationsQuery: () => listState,
    useMarkNotificationReadMutation: () => ({ mutate: mockMutate }),
  };
});

afterEach(() => {
  cleanup();
  mockNavigate.mockReset();
  mockMutate.mockReset();
  listState.data = { items: [], unread_count: 0 };
});

const excluded: NotificationItem = {
  id: 11,
  kind: "promise_excluded",
  payload: {
    kp_id: 42,
    week_start: "2026-09-07",
    reason: "нет арматуры",
  },
  read_at: null,
  created_at: "2026-09-03T12:00:00",
};

const renderBell = () =>
  render(
    <MemoryRouter>
      <NotificationBell />
    </MemoryRouter>,
  );

describe("NotificationBell", () => {
  it("shows unread badge and hides it when count is zero", () => {
    listState.data = { items: [excluded], unread_count: 1 };
    const { rerender } = renderBell();

    expect(screen.getByTestId("notification-badge")).toHaveTextContent("1");

    listState.data = { items: [{ ...excluded, read_at: "2026-09-03T18:00:00" }], unread_count: 0 };
    rerender(
      <MemoryRouter>
        <NotificationBell />
      </MemoryRouter>,
    );
    expect(screen.queryByTestId("notification-badge")).not.toBeInTheDocument();
  });

  it("opens popover with exclusion text and navigates to КП after mark-read", () => {
    listState.data = { items: [excluded], unread_count: 1 };
    renderBell();

    fireEvent.click(screen.getByTestId("notification-bell"));
    expect(screen.getByTestId("notification-popover")).toHaveTextContent(
      "КП №42 снято с недели 7.09: нет арматуры",
    );

    fireEvent.click(screen.getByRole("button", { name: /КП №42/ }));
    expect(mockMutate).toHaveBeenCalledWith(11);
    expect(mockNavigate).toHaveBeenCalledWith("/archive?kp=42");
  });

  it("shows empty state when there are no notifications", () => {
    renderBell();
    fireEvent.click(screen.getByTestId("notification-bell"));
    expect(screen.getByText("Нет уведомлений")).toBeInTheDocument();
  });
});
