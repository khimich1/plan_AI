import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { SubstrateRecommendation } from "@/features/production/types/production";
import { SubstrateRecommendationsBlock } from "./SubstrateRecommendationsBlock";

const baseRec = (
  overrides: Partial<SubstrateRecommendation> = {},
): SubstrateRecommendation => ({
  plate_id: 456,
  kp_id: 127,
  plate_name: "ПБ 57-4,8 ×8п",
  qty_recommended: 3,
  under_plate_id: 123,
  under_kp_id: 115,
  under_plate_name: "ПБ 57-7,2 ×8п",
  needed_by: "2026-09-05",
  storage_days: 24,
  saving_mm: 480,
  saving_m: 2.4,
  ...overrides,
});

describe("SubstrateRecommendationsBlock", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("lists plate_name, qty_recommended, under_plate_name, needed_by, storage_days, saving_m", () => {
    render(
      <SubstrateRecommendationsBlock
        recommendations={[baseRec()]}
        selectedPlatesByKp={{ 127: [456] }}
        onAnalyze={vi.fn()}
        onToggleRecommendation={vi.fn()}
      />,
    );

    expect(screen.getByText("Подложки из поздних КП")).toBeInTheDocument();
    expect(screen.getByText("ПБ 57-4,8 ×8п")).toBeInTheDocument();
    expect(screen.getByText("3")).toBeInTheDocument();
    expect(screen.getByText("ПБ 57-7,2 ×8п")).toBeInTheDocument();
    expect(screen.getByText("05.09")).toBeInTheDocument();
    expect(screen.getByText("24")).toBeInTheDocument();
    expect(screen.getByText("2.4")).toBeInTheDocument();
  });

  it("sorts by saving_m descending", () => {
    render(
      <SubstrateRecommendationsBlock
        recommendations={[
          baseRec({ plate_id: 1, plate_name: "Малая", saving_m: 1.0 }),
          baseRec({ plate_id: 2, plate_name: "Большая", saving_m: 5.5 }),
          baseRec({ plate_id: 3, plate_name: "Средняя", saving_m: 3.2 }),
        ]}
        selectedPlatesByKp={{}}
        onAnalyze={vi.fn()}
        onToggleRecommendation={vi.fn()}
      />,
    );

    const names = screen
      .getAllByRole("row")
      .slice(1)
      .map((row) => within(row).getAllByRole("cell")[1]?.textContent);
    expect(names).toEqual(["Большая", "Средняя", "Малая"]);
  });

  it("checkboxes reflect selection; toggle calls onToggleRecommendation", () => {
    const onToggle = vi.fn();
    const rec = baseRec();
    render(
      <SubstrateRecommendationsBlock
        recommendations={[rec]}
        selectedPlatesByKp={{ 127: [456] }}
        onAnalyze={vi.fn()}
        onToggleRecommendation={onToggle}
      />,
    );

    const checkbox = screen.getByRole("checkbox", { name: "Выбрать ПБ 57-4,8 ×8п" });
    expect(checkbox).toBeChecked();

    fireEvent.click(checkbox);
    expect(onToggle).toHaveBeenCalledWith(rec);
  });

  it("button «Найти подложки» calls onAnalyze", () => {
    const onAnalyze = vi.fn();
    render(
      <SubstrateRecommendationsBlock
        recommendations={[]}
        selectedPlatesByKp={{}}
        onAnalyze={onAnalyze}
        onToggleRecommendation={vi.fn()}
      />,
    );

    fireEvent.click(screen.getByRole("button", { name: "Найти подложки" }));
    expect(onAnalyze).toHaveBeenCalledTimes(1);
  });

  it("shows loading and disclaimer", () => {
    render(
      <SubstrateRecommendationsBlock
        recommendations={[]}
        selectedPlatesByKp={{}}
        loading
        onAnalyze={vi.fn()}
        onToggleRecommendation={vi.fn()}
      />,
    );

    expect(screen.getByText(/Анализируем бэклог/)).toBeInTheDocument();
    expect(
      screen.getByText("Рекомендация — преселектор. Финальный состав может отличаться"),
    ).toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Найти подложки" })).toBeDisabled();
  });

  it("shows error Alert and hides empty-state copy", () => {
    render(
      <SubstrateRecommendationsBlock
        recommendations={[]}
        selectedPlatesByKp={{}}
        errorMessage="Оптимизатор вернул ошибку при анализе подложек: infeasible"
        onAnalyze={vi.fn()}
        onToggleRecommendation={vi.fn()}
      />,
    );

    expect(
      screen.getByText(
        "Оптимизатор вернул ошибку при анализе подложек: infeasible",
      ),
    ).toBeInTheDocument();
    expect(
      screen.queryByText(/Нет рекомендаций по подложкам/),
    ).not.toBeInTheDocument();
  });
});
