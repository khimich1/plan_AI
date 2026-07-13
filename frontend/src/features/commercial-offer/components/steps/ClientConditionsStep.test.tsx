import { cleanup, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ClientConditionsStep } from "@/features/commercial-offer/components/steps/ClientConditionsStep";

const mockUseAuth = vi.fn();

vi.mock("@/features/auth/model/AuthProvider", () => ({
  useAuth: () => mockUseAuth(),
}));

const managers = [
  { id: 10, fio: "Иванов И.И.", contact_number: "+7 900 000-00-01", email: "ivanov@example.com" },
  { id: 20, fio: "Петров П.П.", contact_number: "+7 900 000-00-02", email: "petrov@example.com" },
];

const defaultProps = {
  managers,
  selectedManagerId: null as number | null,
  defaultValues: {
    clientName: "",
    conditionsMode: "standard" as const,
    deliveryConditions: "",
    paymentConditions: "",
  },
  errorMessage: null,
  isPending: false,
  onBack: vi.fn(),
  onManagerChange: vi.fn(),
  onSubmit: vi.fn(),
};

describe("ClientConditionsStep", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
  });

  it("auto-selects manager from auth profile when manager_id is in list", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, username: "manager", role: "manager", manager_id: 10, is_active: true },
    });

    const onManagerChange = vi.fn();
    render(<ClientConditionsStep {...defaultProps} onManagerChange={onManagerChange} />);

    expect(onManagerChange).toHaveBeenCalledWith(10);
    expect(screen.getByText("Иванов И.И.")).toBeInTheDocument();
  });

  it("shows manager select when profile has no manager_id", () => {
    mockUseAuth.mockReturnValue({
      user: { id: 1, username: "admin", role: "admin", manager_id: null, is_active: true },
    });

    render(<ClientConditionsStep {...defaultProps} />);

    expect(screen.getByText("Выберите менеджера для итоговых документов КП.")).toBeInTheDocument();
    expect(screen.getByRole("combobox")).toBeInTheDocument();
  });
});
