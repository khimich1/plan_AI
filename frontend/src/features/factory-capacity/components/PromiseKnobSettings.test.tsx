import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import type { ReactNode } from "react";
import { PromiseKnobSettings } from "@/features/factory-capacity/components/PromiseKnobSettings";

const mutateAsync = vi.fn();
const mutationReset = vi.fn();
const mutationState = {
  isPending: false,
  isError: false,
  error: null as Error | null,
};

vi.mock("@/features/factory-capacity/api/promiseQuote", async () => {
  const actual = await vi.importActual<typeof import("@/features/factory-capacity/api/promiseQuote")>(
    "@/features/factory-capacity/api/promiseQuote",
  );
  return {
    ...actual,
    useUpdatePromiseKnobMutation: () => ({
      mutateAsync,
      isPending: mutationState.isPending,
      isError: mutationState.isError,
      error: mutationState.error,
      reset: mutationReset,
    }),
  };
});

afterEach(() => {
  cleanup();
  mutateAsync.mockReset();
  mutationReset.mockReset();
  mutationState.isPending = false;
  mutationState.isError = false;
  mutationState.error = null;
});

const wrap = (ui: ReactNode) => {
  const client = new QueryClient({ defaultOptions: { queries: { retry: false } } });
  return render(<QueryClientProvider client={client}>{ui}</QueryClientProvider>);
};

describe("PromiseKnobSettings", () => {
  it("shows current knob and opens an inline confirm form", () => {
    wrap(<PromiseKnobSettings currentKnob={3} />);

    fireEvent.click(screen.getByRole("button", { name: /Настроить ручку/i }));

    expect(screen.getByLabelText("Дорожек в день")).toHaveValue(3);
    expect(screen.getByText("Влияет только на новые расчёты")).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /Подтвердить/i })).toBeDisabled();
  });

  it("confirms a new value in range 1..5", async () => {
    mutateAsync.mockResolvedValue({
      tracks_per_day: 4,
      updated_by: "tester",
      updated_at: "2026-09-03T15:30:00",
      min: 1,
      max: 5,
    });

    wrap(<PromiseKnobSettings currentKnob={3} />);
    fireEvent.click(screen.getByRole("button", { name: /Настроить ручку/i }));
    fireEvent.change(screen.getByLabelText("Дорожек в день"), { target: { value: "4" } });

    const confirm = screen.getByRole("button", { name: /Подтвердить/i });
    expect(confirm).toBeEnabled();
    fireEvent.click(confirm);

    expect(mutateAsync).toHaveBeenCalledWith(4);
  });

  it("keeps confirm disabled for 0 and 6", () => {
    wrap(<PromiseKnobSettings currentKnob={3} />);
    fireEvent.click(screen.getByRole("button", { name: /Настроить ручку/i }));

    fireEvent.change(screen.getByLabelText("Дорожек в день"), { target: { value: "0" } });
    expect(screen.getByRole("button", { name: /Подтвердить/i })).toBeDisabled();

    fireEvent.change(screen.getByLabelText("Дорожек в день"), { target: { value: "6" } });
    expect(screen.getByRole("button", { name: /Подтвердить/i })).toBeDisabled();
    expect(mutateAsync).not.toHaveBeenCalled();
  });
});
