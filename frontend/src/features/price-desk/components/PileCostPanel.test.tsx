import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { pileCostApi } from "@/features/price-desk/api/pileCostApi";
import { PileCostPanel } from "@/features/price-desk/components/PileCostPanel";
import type {
  PileCostInputsBody,
  PileCostMarkBreakdown,
  PileCostPreview,
  PileCostState,
} from "@/features/price-desk/types/pileCost";
import { ApiError } from "@/shared/lib/apiError";

vi.mock("@/features/price-desk/api/pileCostApi", () => ({
  pileCostApi: {
    get: vi.fn(),
    save: vi.fn(),
    preview: vi.fn(),
    apply: vi.fn(),
  },
}));

const GRADES = ["B15", "B20", "B22_5", "B25", "B30_granite"] as const;

let cementPrice = 7.54;
let buyerB15 = 3513.2818728;

const inputs = (cement: number): PileCostInputsBody => ({
  prices: [
    { article: "cement", price: cement },
    { article: "crushed_20", price: 1.8 },
    { article: "crushed_5", price: 1.97 },
    { article: "sand", price: 0.66 },
    { article: "additive", price: 59 },
    { article: "wire_5", price: 72.7 },
    { article: "ai_10", price: 50.57 },
    { article: "a500c_10", price: 51.39 },
    { article: "a500c_12", price: 51.39 },
  ],
  recipes: GRADES.map((grade, index) => ({
    concrete_grade: grade,
    cement_kg_m3: 320 + index * 10,
    sand_kg_m3: 990,
    crushed_kg_m3: 950,
  })),
  overhead: { general: 3186, payroll: 3400, forming: 40, other: 606 },
});

const breakdown = (buyer: number): PileCostMarkBreakdown => ({
  mark: "С30.20-3",
  section_cm: 20,
  volume_m3: 0.13,
  grades: GRADES.map((grade) => ({
    concrete_grade: grade,
    materials_rub: 1431.4179,
    expenses_rub: 968.3648,
    cost: 2399.7827,
    buyer_price: grade === "B15" ? buyer : 100,
    surcharge_applied: false,
    cement_kg: 41.6,
    sand_kg: 128.7,
    crushed_kg: 123.5,
    additive_kg: 0.2496,
    cement_rub: 313.664,
    sand_rub: 84.942,
    crushed_rub: 222.3,
    additive_rub: 14.7264,
    crushed_article: "crushed_20",
    lines: [{ article: "a500c_12", kg: 8, rub: 411.12 }],
  })),
});

const stateFor = (mark?: string | null): PileCostState => ({
  inputs_sha256: cementPrice === 7.54 ? "sha-1" : "sha-2",
  inputs: inputs(cementPrice),
  marks: [
    {
      mark: "С30.20-3",
      section_cm: 20,
      grades: GRADES.map((grade) => ({ concrete_grade: grade, buyer_price: grade === "B15" ? buyerB15 : 100 })),
    },
    {
      mark: "С40.30-1",
      section_cm: 30,
      grades: GRADES.map((grade) => ({ concrete_grade: grade, buyer_price: 9033.108758784 })),
    },
  ],
  breakdown: mark ? breakdown(buyerB15) : null,
});

const PREVIEW: PileCostPreview = {
  inputs_sha256: "preview-sha",
  changed: 2,
  new: 1,
  missing: 1,
  unchanged: 4,
  examples: {
    changed: [{ key: "С30.20-3 / B15", old_price: 1000, new_price: 3513.2818728 }],
    new: [],
    missing: [{ key: "С999 / B15", old_price: 50.5, new_price: null }],
    unchanged: [],
  },
};

const renderPanel = () => {
  const client = new QueryClient({
    defaultOptions: { queries: { retry: false }, mutations: { retry: false } },
  });
  return render(
    <QueryClientProvider client={client}>
      <PileCostPanel />
    </QueryClientProvider>,
  );
};

const openMark = async () => {
  fireEvent.change(screen.getByLabelText("Найти марку"), { target: { value: "С30.20-3" } });
  fireEvent.click(screen.getByRole("button", { name: "С30.20-3" }));
  return screen.findByRole("region", { name: "Разложение С30.20-3" });
};

beforeEach(() => {
  cementPrice = 7.54;
  buyerB15 = 3513.2818728;
  vi.mocked(pileCostApi.get).mockImplementation(async (mark) => stateFor(mark));
  vi.mocked(pileCostApi.save).mockImplementation(async (body) => {
    const next = body.prices.find((row) => row.article === "cement")?.price;
    if (typeof next === "number") {
      cementPrice = next;
    }
    buyerB15 = 3700.5;
    return { inputs_sha256: "sha-2", inputs: body };
  });
});

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe("PileCostPanel", () => {
  it("updates the mark card when the cement price is saved and does not apply", async () => {
    renderPanel();
    expect(await screen.findByLabelText("Цена: Цемент")).toHaveValue("7.54");
    const card = await openMark();

    expect(card.textContent).toContain("3513,2818728");
    expect(card.textContent).toContain("1431,4179");
    expect(card.textContent).toContain("Объём 0,13 м³");
    expect(card.textContent).toContain("А500С Ø12: 8 кг");
    expect(card.querySelector("input, textarea")).toBeNull();
    expect(screen.queryByRole("textbox", { name: /объём/i })).not.toBeInTheDocument();
    expect(screen.queryByRole("textbox", { name: /арматур/i })).not.toBeInTheDocument();
    expect(screen.getByText(/надбавка к расходам 3%/)).toBeInTheDocument();
    expect(screen.getByText(/×1,2/)).toBeInTheDocument();
    expect(screen.getByText(/×1,22/)).toBeInTheDocument();
    expect(screen.getByText(/×1,02/)).toBeInTheDocument();

    fireEvent.change(screen.getByLabelText("Цена: Цемент"), { target: { value: "8.2" } });
    fireEvent.blur(screen.getByLabelText("Цена: Цемент"));

    await waitFor(() => {
      expect(pileCostApi.save).toHaveBeenCalled();
    });
    const sent = vi.mocked(pileCostApi.save).mock.calls[0][0];
    expect(sent.prices.find((row) => row.article === "cement")?.price).toBe(8.2);
    expect(sent.prices.find((row) => row.article === "a500c_12")?.price).toBe(51.39);
    expect(Object.keys(sent).sort()).toEqual(["overhead", "prices", "recipes"]);
    expect(JSON.stringify(sent)).not.toContain("volume_m3");
    expect(JSON.stringify(sent)).not.toMatch(/"kg"/);

    await waitFor(() => {
      expect(screen.getByRole("region", { name: "Разложение С30.20-3" }).textContent).toContain("3700,5");
    });
    expect(pileCostApi.preview).not.toHaveBeenCalled();
    expect(pileCostApi.apply).not.toHaveBeenCalled();
  });

  it("does not apply until preview, then sends the preview sha", async () => {
    vi.mocked(pileCostApi.preview).mockResolvedValue(PREVIEW);
    vi.mocked(pileCostApi.apply).mockResolvedValue({ ...PREVIEW, changed: 0, new: 0 });
    renderPanel();
    const panel = await screen.findByRole("region", { name: "Цельные сваи" });
    const write = await within(panel).findByRole("button", { name: "Записать в прайс" });
    expect(write).toBeDisabled();
    fireEvent.click(write);
    expect(pileCostApi.apply).not.toHaveBeenCalled();
    expect(pileCostApi.preview).not.toHaveBeenCalled();

    fireEvent.click(within(panel).getByRole("button", { name: "Показать, что изменится" }));
    expect(await within(panel).findByText(/3513,2818728/)).toBeInTheDocument();
    expect(within(panel).getByText(/нет в норме 1/)).toBeInTheDocument();
    expect(pileCostApi.preview).toHaveBeenCalledTimes(1);
    expect(pileCostApi.preview).toHaveBeenCalledWith();
    expect(pileCostApi.apply).not.toHaveBeenCalled();

    fireEvent.click(within(panel).getByRole("button", { name: "Отмена" }));
    expect(within(panel).getByRole("button", { name: "Записать в прайс" })).toBeDisabled();
    expect(pileCostApi.apply).not.toHaveBeenCalled();

    fireEvent.click(within(panel).getByRole("button", { name: "Показать, что изменится" }));
    await within(panel).findByText(/3513,2818728/);
    fireEvent.click(within(panel).getByRole("button", { name: "Записать в прайс" }));
    await waitFor(() => {
      expect(pileCostApi.apply).toHaveBeenCalledTimes(1);
    });
    expect(pileCostApi.apply).toHaveBeenCalledWith({ inputs_sha256: "preview-sha" });
    expect(await within(panel).findByText("Записано в прайс.")).toBeInTheDocument();
  });

  it("shows a missing article price and does not apply", async () => {
    vi.mocked(pileCostApi.preview).mockRejectedValue(new ApiError("нет цены статьи wire_5", 422));
    renderPanel();
    const panel = await screen.findByRole("region", { name: "Цельные сваи" });
    fireEvent.click(await within(panel).findByRole("button", { name: "Показать, что изменится" }));
    expect(await screen.findByRole("alert")).toHaveTextContent("нет цены статьи wire_5");
    expect(pileCostApi.apply).not.toHaveBeenCalled();
    expect(within(panel).getByRole("button", { name: "Записать в прайс" })).toBeDisabled();
  });

  it("shows a fingerprint mismatch and keeps the diff", async () => {
    vi.mocked(pileCostApi.preview).mockResolvedValue(PREVIEW);
    vi.mocked(pileCostApi.apply).mockRejectedValue(
      new ApiError("расчёт изменился, посмотрите дифф ещё раз", 409),
    );
    renderPanel();
    const panel = await screen.findByRole("region", { name: "Цельные сваи" });
    fireEvent.click(await within(panel).findByRole("button", { name: "Показать, что изменится" }));
    await within(panel).findByText(/3513,2818728/);
    fireEvent.click(within(panel).getByRole("button", { name: "Записать в прайс" }));
    expect(await screen.findByRole("alert")).toHaveTextContent("расчёт изменился, посмотрите дифф ещё раз");
    expect(pileCostApi.apply).toHaveBeenCalledTimes(1);
    expect(pileCostApi.apply).toHaveBeenCalledWith({ inputs_sha256: "preview-sha" });
    expect(within(panel).getByText(/3513,2818728/)).toBeInTheDocument();
  });

  it("rejects a non-positive price before save", async () => {
    renderPanel();
    const cement = await screen.findByLabelText("Цена: Цемент");
    fireEvent.change(cement, { target: { value: "0" } });
    fireEvent.blur(cement);
    expect(await screen.findByRole("alert")).toHaveTextContent("больше нуля");
    expect(pileCostApi.save).not.toHaveBeenCalled();
    expect(pileCostApi.apply).not.toHaveBeenCalled();
  });

  it("says when the mark is not in the norm", async () => {
    renderPanel();
    const search = await screen.findByLabelText("Найти марку");
    fireEvent.change(search, { target: { value: "С130.35-13А**" } });
    expect(await screen.findByText("Марка не найдена")).toBeInTheDocument();
    expect(pileCostApi.get).not.toHaveBeenCalledWith("С130.35-13А**");
  });
});
