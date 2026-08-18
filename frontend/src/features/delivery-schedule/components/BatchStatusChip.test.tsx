import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it } from "vitest";
import { BatchStatusChip } from "@/features/delivery-schedule/components/BatchStatusChip";

describe("BatchStatusChip", () => {
  afterEach(() => {
    cleanup();
  });

  it("renders green/yellow/red labels with hint", () => {
    const { rerender } = render(<BatchStatusChip status="green" hint="запас 5 дн" />);
    expect(screen.getByText("В срок")).toBeInTheDocument();
    expect(screen.getByText(/запас 5 дн/)).toBeInTheDocument();

    rerender(<BatchStatusChip status="yellow" />);
    expect(screen.getByText("На грани")).toBeInTheDocument();

    rerender(<BatchStatusChip status="red" hint="нужно +2 дорожки" />);
    expect(screen.getByText("Риск срыва")).toBeInTheDocument();
    expect(screen.getByText(/нужно \+2 дорожки/)).toBeInTheDocument();
  });

  it("shows placeholder when status is null", () => {
    render(<BatchStatusChip status={null} />);
    expect(screen.getByText("нет статуса")).toBeInTheDocument();
  });
});
