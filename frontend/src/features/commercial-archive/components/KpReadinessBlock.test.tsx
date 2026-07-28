import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { KpReadinessBlock } from "@/features/commercial-archive/components/KpReadinessBlock";
import type { KpReadinessSummary } from "@/features/commercial-archive/types/archive";

const mockUseKpReadinessPositionsQuery = vi.fn();

vi.mock("@/features/commercial-archive/hooks/useArchiveQueries", () => ({
  useKpReadinessPositionsQuery: (...args: unknown[]) => mockUseKpReadinessPositionsQuery(...args),
}));

function makeReadiness(overrides: Partial<KpReadinessSummary> = {}): KpReadinessSummary {
  return {
    completion_percentage: 72,
    sgp_progress: { n: 14, m: 20 },
    issuable_qty: 14,
    in_production_qty: 6,
    summary_text: "14 из 20 шт на складе, 6 в производстве. Можно выдать 14 шт.",
    client_copy_text: "Здравствуйте! По вашему заказу №42: 14 из 20 шт уже на складе.",
    steps: [
      { id: "kp", label: "КП", state: "done" },
      { id: "production", label: "Производство", state: "active", hint: "72%" },
      { id: "sgp", label: "СГП", state: "active", hint: "14/20" },
      { id: "release", label: "Выдача", state: "disabled" },
      { id: "closed", label: "Закрыто", state: "disabled" },
    ],
    release_note: "Выдача с СГП — в следующем обновлении",
    ...overrides,
  };
}

describe("KpReadinessBlock", () => {
  beforeEach(() => {
    mockUseKpReadinessPositionsQuery.mockReturnValue({
      data: undefined,
      isPending: false,
      isError: false,
      error: null,
    });
    Object.assign(navigator, {
      clipboard: { writeText: vi.fn().mockResolvedValue(undefined) },
    });
  });

  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("renders stepper, summary and release note", () => {
    render(<KpReadinessBlock kpId={42} readiness={makeReadiness()} />);

    expect(screen.getByText("Статус производства")).toBeInTheDocument();
    expect(screen.getByText(/14 из 20 шт на складе/)).toBeInTheDocument();
    expect(screen.getByText("Выдача с СГП — в следующем обновлении")).toBeInTheDocument();
    expect(screen.getByText("Производство")).toBeInTheDocument();
    expect(screen.getByText("СГП")).toBeInTheDocument();
    expect(screen.getByText("Выдача")).toBeInTheDocument();
  });

  it("does not fetch positions until expanded", () => {
    render(<KpReadinessBlock kpId={42} readiness={makeReadiness()} />);

    expect(mockUseKpReadinessPositionsQuery).toHaveBeenCalledWith(42, { enabled: false });
  });

  it("fetches positions when expanded", async () => {
    mockUseKpReadinessPositionsQuery.mockImplementation((_kpId, opts) => ({
      data: opts?.enabled
        ? {
            items: [
              {
                position_number: 1,
                plate_name: "ПБ 59-12-8",
                length_m: 5.9,
                width_m: 1.2,
                load_class: 800,
                label: "ПБ 59-12-8",
                ordered: 10,
                in_plan: 4,
                on_sgp: 6,
                remaining: 0,
              },
            ],
            count: 1,
          }
        : undefined,
      isPending: false,
      isError: false,
      error: null,
    }));

    render(<KpReadinessBlock kpId={42} readiness={makeReadiness()} />);

    fireEvent.click(screen.getByRole("button", { name: /Подробнее/i }));

    await waitFor(() => {
      expect(mockUseKpReadinessPositionsQuery).toHaveBeenLastCalledWith(42, { enabled: true });
    });
    expect(screen.getByText("ПБ 59-12-8")).toBeInTheDocument();
    expect(screen.getByText("10")).toBeInTheDocument();
  });

  it("copies client_copy_text to clipboard", async () => {
    render(<KpReadinessBlock kpId={42} readiness={makeReadiness()} />);

    fireEvent.click(screen.getByRole("button", { name: /Скопировать для клиента/i }));

    await waitFor(() => {
      expect(navigator.clipboard.writeText).toHaveBeenCalledWith(
        "Здравствуйте! По вашему заказу №42: 14 из 20 шт уже на складе.",
      );
    });
    expect(screen.getByText(/скопирован/i)).toBeInTheDocument();
  });

  it("renders expected SGP date line when label is present", () => {
    render(
      <KpReadinessBlock
        kpId={42}
        readiness={makeReadiness({
          expected_sgp_date: "2026-08-14",
          expected_sgp_date_label: "14.08.2026",
          fully_scheduled: true,
        })}
      />,
    );

    expect(screen.getByText(/Ожидаем на СГП к:/)).toBeInTheDocument();
    expect(screen.getByText("14.08.2026")).toBeInTheDocument();
  });

  it("does not render expected SGP date line when label is absent", () => {
    render(<KpReadinessBlock kpId={42} readiness={makeReadiness()} />);

    expect(screen.queryByText(/Ожидаем на СГП к:/)).not.toBeInTheDocument();
  });
});
