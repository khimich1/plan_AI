import { cleanup, fireEvent, render, screen, within } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import type {
  PendingPromiseExclusion,
  PromisedBlockItem,
} from "@/features/production/types/production";
import { PromisedWeekBlock } from "./PromisedWeekBlock";

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

const activeItem: PromisedBlockItem = {
  kp_id: 88,
  promised_date: "2026-09-25",
  tracks: 2,
  status: "active",
  week_start: "2026-06-15",
  customer_name: "Обещанный",
};

const overdueItem: PromisedBlockItem = {
  kp_id: 91,
  promised_date: "2026-09-04",
  tracks: 3,
  status: "overdue",
  week_start: "2026-06-08",
  customer_name: "Просроченный",
};

const pendingWhole: PendingPromiseExclusion = {
  kpId: 88,
  weekStart: "2026-06-15",
  kind: "whole",
};

describe("PromisedWeekBlock", () => {
  it("shows promised KPs as preselected checkboxes", () => {
    render(
      <PromisedWeekBlock
        items={[activeItem]}
        selectedPlatesByKp={{ 88: [501] }}
        pendingExclusion={null}
        onToggleKp={vi.fn()}
        onConfirmExclusion={vi.fn()}
        onCancelExclusion={vi.fn()}
      />,
    );

    expect(screen.getByText("Обещано на эту неделю")).toBeInTheDocument();
    expect(screen.getByText("КП №88")).toBeInTheDocument();
    expect(screen.getByText("Обещанный")).toBeInTheDocument();
    expect(screen.getByText("2 дор.")).toBeInTheDocument();
    expect(screen.getByLabelText("Снять обещанное КП 88")).toBeChecked();
  });

  it("renders overdue items in a distinct red block at the top", () => {
    render(
      <PromisedWeekBlock
        items={[activeItem, overdueItem]}
        selectedPlatesByKp={{ 88: [501], 91: [601] }}
        pendingExclusion={null}
        onToggleKp={vi.fn()}
        onConfirmExclusion={vi.fn()}
        onCancelExclusion={vi.fn()}
      />,
    );

    const overdue = screen.getByTestId("promised-overdue-block");
    expect(overdue).toHaveTextContent("Обещано, но не в плане");
    expect(overdue).toHaveTextContent("КП №91");
    expect(overdue).toHaveTextContent("просрочено");
    expect(within(overdue).queryByText("КП №88")).not.toBeInTheDocument();

    const active = screen.getByTestId("promised-active-block");
    expect(active).toHaveTextContent("КП №88");
    expect(within(active).queryByText("КП №91")).not.toBeInTheDocument();
  });

  it("asks for a reason instead of unchecking when the box is clicked", () => {
    const onToggleKp = vi.fn();
    render(
      <PromisedWeekBlock
        items={[activeItem]}
        selectedPlatesByKp={{ 88: [501] }}
        pendingExclusion={null}
        onToggleKp={onToggleKp}
        onConfirmExclusion={vi.fn()}
        onCancelExclusion={vi.fn()}
      />,
    );

    fireEvent.click(screen.getByLabelText("Снять обещанное КП 88"));
    expect(onToggleKp).toHaveBeenCalledWith(88);
  });

  it("does not confirm an empty reason", () => {
    const onConfirmExclusion = vi.fn();
    render(
      <PromisedWeekBlock
        items={[activeItem]}
        selectedPlatesByKp={{ 88: [501] }}
        pendingExclusion={pendingWhole}
        onToggleKp={vi.fn()}
        onConfirmExclusion={onConfirmExclusion}
        onCancelExclusion={vi.fn()}
      />,
    );

    const confirm = screen.getByRole("button", { name: "Подтвердить причину" });
    expect(confirm).toBeDisabled();
    fireEvent.change(screen.getByLabelText("Причина исключения обещанного КП"), {
      target: { value: "   " },
    });
    expect(confirm).toBeDisabled();
    fireEvent.click(confirm);
    expect(onConfirmExclusion).not.toHaveBeenCalled();
  });

  it("submits a non-empty reason", () => {
    const onConfirmExclusion = vi.fn();
    render(
      <PromisedWeekBlock
        items={[activeItem]}
        selectedPlatesByKp={{ 88: [501] }}
        pendingExclusion={pendingWhole}
        onToggleKp={vi.fn()}
        onConfirmExclusion={onConfirmExclusion}
        onCancelExclusion={vi.fn()}
      />,
    );

    fireEvent.change(screen.getByLabelText("Причина исключения обещанного КП"), {
      target: { value: "Клиент перенёс поставку" },
    });
    fireEvent.click(screen.getByRole("button", { name: "Подтвердить причину" }));
    expect(onConfirmExclusion).toHaveBeenCalledWith("Клиент перенёс поставку");
  });

  it("shows an empty state when there are no promised KPs", () => {
    render(
      <PromisedWeekBlock
        items={[]}
        selectedPlatesByKp={{}}
        pendingExclusion={null}
        onToggleKp={vi.fn()}
        onConfirmExclusion={vi.fn()}
        onCancelExclusion={vi.fn()}
      />,
    );

    expect(screen.getByTestId("promised-week-block")).toHaveTextContent(
      "Нет обещанных КП на выбранные дни.",
    );
  });
});
