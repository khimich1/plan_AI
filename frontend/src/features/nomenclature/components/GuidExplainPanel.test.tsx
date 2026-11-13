import { cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { GuidExplainPanel } from "@/features/nomenclature/components/GuidExplainPanel";
import type {
  GuidExplainResponse,
  GuidExplainSource,
  GuidSearchResponse,
} from "@/features/nomenclature/types/nomenclature";

const explainState: { data?: GuidExplainResponse; isError: boolean; isLoading: boolean } = {
  isError: false,
  isLoading: false,
};
const searchState: { data?: GuidSearchResponse } = {};
const explainSpy = vi.fn();
const searchSpy = vi.fn();

vi.mock("@/features/nomenclature/hooks/useGuidExplainQueries", () => ({
  useGuidExplainQuery: (...args: unknown[]) => {
    explainSpy(...args);
    return explainState;
  },
  useGuidSearchQuery: (...args: unknown[]) => {
    searchSpy(...args);
    return { data: searchState.data, isError: false, isLoading: false };
  },
}));

const SOURCE: GuidExplainSource = {
  scope: "test",
  explain: vi.fn(),
  search: vi.fn(),
};

const EXPLAIN: GuidExplainResponse = {
  product_kind: "plate",
  mark: "Плиты ПБ 22,80-12-8п",
  reason: "нет GUID 1С",
  hint: "заведите карточку в 1С и загрузите отчёт",
  same_cards: [
    {
      mark: "Плиты ПБ 22,8-12-8п",
      guid_count: 1,
      label: "лишний ноль",
      selectable: true,
      cards: [],
    },
  ],
  slice: [
    {
      mark: "Плиты ПБ 22,8-10,2-8п",
      guid_count: 2,
      label: null,
      selectable: false,
      cards: [],
    },
    { mark: "Плиты ПБ 22,8-9-8п", guid_count: 0, label: null, selectable: false, cards: [] },
  ],
};

const renderPanel = () =>
  render(
    <GuidExplainPanel source={SOURCE} productKind="plate" mark="Плиты ПБ 22,80-12-8п" />,
  );

const open = () =>
  fireEvent.click(screen.getByRole("button", { name: /Почему нет номера/ }));

describe("GuidExplainPanel", () => {
  afterEach(() => {
    cleanup();
    vi.clearAllMocks();
    explainState.data = undefined;
    explainState.isError = false;
    explainState.isLoading = false;
    searchState.data = undefined;
  });

  it("stays collapsed and does not load anything until opened", () => {
    explainState.data = EXPLAIN;
    renderPanel();

    expect(screen.getByRole("button", { name: /Почему нет номера/ })).toHaveAttribute(
      "aria-expanded",
      "false",
    );
    expect(explainSpy).not.toHaveBeenCalled();
    expect(screen.queryByText(/нет GUID 1С/)).not.toBeInTheDocument();
  });

  it("shows the reason, the hint and the same card with the number count", () => {
    explainState.data = EXPLAIN;
    renderPanel();
    open();

    expect(screen.getByText(/нет GUID 1С/)).toBeInTheDocument();
    expect(screen.getByText(/заведите карточку в 1С/)).toBeInTheDocument();
    const same = screen.getByTestId("guid-explain-same-cards");
    expect(within(same).getByText(/Плиты ПБ 22,8-12-8п/)).toBeInTheDocument();
    expect(within(same).getByText(/номеров: 1/)).toBeInTheDocument();
    expect(within(same).getByText(/лишний ноль/)).toBeInTheDocument();
  });

  it("lists the slice and marks rows without numbers", () => {
    explainState.data = EXPLAIN;
    renderPanel();
    open();

    const slice = screen.getByTestId("guid-explain-slice");
    expect(within(slice).getAllByRole("listitem")).toHaveLength(2);
    expect(within(slice).getByText(/номеров: 2/)).toBeInTheDocument();
    expect(within(slice).getByText(/номера нет/)).toBeInTheDocument();
  });

  it("says so when the line already has a number", () => {
    explainState.data = { ...EXPLAIN, reason: "", hint: "", same_cards: [], slice: [] };
    renderPanel();
    open();

    expect(screen.getByText(/Номер 1С у этой строки есть/)).toBeInTheDocument();
  });

  it("shows an error when the explanation cannot be loaded", () => {
    explainState.isError = true;
    renderPanel();
    open();

    expect(screen.getByRole("alert")).toHaveTextContent(/Не удалось загрузить/);
  });

  it("asks for three characters and searches only after the debounce", async () => {
    explainState.data = EXPLAIN;
    searchState.data = {
      items: [
        { mark: "Плиты ПБ 30-12-8п", guid_count: 0, label: null, selectable: false, cards: [] },
      ],
    };
    renderPanel();
    open();

    const input = screen.getByLabelText("Найти марку");
    fireEvent.change(input, { target: { value: "ПБ" } });
    expect(screen.getByText(/не меньше 3 знаков/)).toBeInTheDocument();
    expect(screen.queryByTestId("guid-explain-search-results")).not.toBeInTheDocument();

    fireEvent.change(input, { target: { value: "ПБ 30" } });
    await waitFor(() => {
      expect(searchSpy).toHaveBeenLastCalledWith(SOURCE, "plate", "ПБ 30");
    });
    const results = await screen.findByTestId("guid-explain-search-results");
    expect(within(results).getByText(/Плиты ПБ 30-12-8п/)).toBeInTheDocument();
    expect(within(results).getByText(/номера нет/)).toBeInTheDocument();
  });

  it("offers replace only on a selectable row and does not write without confirmation", async () => {
    const preview = vi.fn().mockResolvedValue({
      items: [
        {
          kp_id: 47,
          qty: 2,
          was: money(100),
          will: money(9000),
          delivery_was: null,
          delivery_will: null,
          oneoff_notice: "Договорная цена будет заменена ценой прайса.",
        },
      ],
      skipped: [],
    });
    const commit = vi.fn();
    explainState.data = {
      ...EXPLAIN,
      slice: [
        { mark: "Плиты ПБ 22,8-10,2-8п", guid_count: 2, label: null, selectable: true, cards: [] },
        { mark: "Плиты ПБ 22,8-9-8п", guid_count: 0, label: null, selectable: false, cards: [] },
      ],
    };
    render(
      <GuidExplainPanel
        source={{ ...SOURCE, scope: "economist", previewReplace: preview, commitReplace: commit }}
        productKind="plate"
        mark="Плиты ПБ 22,80-12-8п"
        onCommitted={vi.fn()}
      />,
    );
    open();

    const slice = screen.getByTestId("guid-explain-slice");
    expect(within(slice).getAllByRole("button", { name: "Заменить" })).toHaveLength(1);
    fireEvent.click(within(screen.getByTestId("guid-explain-same-cards")).getByRole("button", { name: "Заменить" }));

    expect(await screen.findByText(/Договорная цена будет заменена ценой прайса/)).toBeInTheDocument();
    expect(screen.getByText(/Цена за штуку со скидкой/)).toBeInTheDocument();
    expect(screen.getByText(/Итог КП/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Отмена" }));
    expect(commit).not.toHaveBeenCalled();
  });

  it("writes the replacement only after confirmation", async () => {
    const preview = vi.fn().mockResolvedValue({
      items: [
        {
          kp_id: 47,
          qty: 2,
          was: money(100),
          will: money(9000),
          delivery_was: 1000,
          delivery_will: 2000,
          oneoff_notice: null,
        },
      ],
      skipped: [{ kp_id: 48, reason: "на согласовании" }],
    });
    const commit = vi.fn().mockResolvedValue({
      written: [{ kp_id: 47, total: 18000 }],
      skipped: [],
      queue_notice: "Очередь обновится при открытии страницы.",
    });
    const committed = vi.fn();
    explainState.data = EXPLAIN;
    render(
      <GuidExplainPanel
        source={{ ...SOURCE, scope: "economist", previewReplace: preview, commitReplace: commit }}
        productKind="plate"
        mark="Плиты ПБ 22,80-12-8п"
        onCommitted={committed}
      />,
    );
    open();
    fireEvent.click(within(screen.getByTestId("guid-explain-same-cards")).getByRole("button", { name: "Заменить" }));
    expect(await screen.findByText(/Доставка/)).toBeInTheDocument();
    expect(screen.getByText(/КП №48: на согласовании/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Записать замену" }));
    await waitFor(() => expect(commit).toHaveBeenCalledTimes(1));
    expect(commit.mock.calls[0][0].target_mark).toBe("Плиты ПБ 22,8-12-8п");
    expect(commit.mock.calls[0][0].replace_oneoff).toBe(false);
    expect(committed).toHaveBeenCalled();
    expect(await screen.findByText(/Очередь обновится при открытии страницы/)).toBeInTheDocument();
  });

  it("lets a manager replace only a row with one number", async () => {
    const preview = vi.fn().mockResolvedValue({
      items: [
        {
          kp_id: 0,
          qty: 2,
          was: money(100),
          will: money(200),
          delivery_was: null,
          delivery_will: null,
          oneoff_notice: null,
        },
      ],
      skipped: [],
    });
    const commit = vi.fn().mockResolvedValue({ written: [], skipped: [], queue_notice: null });
    explainState.data = {
      ...EXPLAIN,
      same_cards: [
        { mark: "С110.30-9у", guid_count: 1, label: null, selectable: true, cards: [] },
        { mark: "С110.30-9", guid_count: 0, label: null, selectable: false, cards: [] },
      ],
      slice: [
        { mark: "С9-35T7", guid_count: 0, label: null, selectable: false, cards: [] },
        {
          mark: "С15-35T7",
          guid_count: 2,
          label: "карточек две",
          selectable: false,
          cards: [
            { name: "Свая С15-35T7 завод", guid: "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa" },
            { name: "Свая С15-35T7 особая", guid: "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb" },
          ],
        },
      ],
    };
    render(
      <GuidExplainPanel
        source={{ ...SOURCE, scope: "draft-1", previewReplace: preview, commitReplace: commit }}
        productKind="pile"
        mark="C110.30-9"
      />,
    );
    open();

    expect(screen.getAllByRole("button", { name: "Заменить" })).toHaveLength(1);
    expect(screen.getByText(/С9-35T7/)).toBeInTheDocument();
    expect(screen.getByText(/карточек две/)).toBeInTheDocument();
    expect(screen.queryByRole("button", { name: "Свая С15-35T7 завод" })).not.toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Заменить" }));
    expect(await screen.findByText(/Количество 2/)).toBeInTheDocument();
    fireEvent.click(screen.getByRole("button", { name: "Отмена" }));
    expect(commit).not.toHaveBeenCalled();
  });

  it("lets an economist pick one of two same-named cards without typing a GUID", async () => {
    const preview = vi.fn().mockResolvedValue({
      items: [
        {
          kp_id: 47,
          qty: 2,
          was: money(100),
          will: money(200),
          delivery_was: null,
          delivery_will: null,
          oneoff_notice: null,
        },
      ],
      skipped: [],
    });
    const commit = vi.fn().mockResolvedValue({ written: [], skipped: [], queue_notice: null });
    const name = "Плиты ПБ 22,8-10,2-8п";
    const first = "11111111-1111-1111-1111-111111111111";
    const second = "22222222-2222-2222-2222-222222222222";
    explainState.data = {
      ...EXPLAIN,
      same_cards: [],
      slice: [
        {
          mark: name,
          guid_count: 2,
          label: null,
          selectable: true,
          cards: [
            { name, guid: first },
            { name, guid: second },
          ],
        },
      ],
    };
    render(
      <GuidExplainPanel
        source={{ ...SOURCE, scope: "economist", previewReplace: preview, commitReplace: commit }}
        productKind="plate"
        mark="Плиты ПБ 22,80-12-8п"
      />,
    );
    open();
    fireEvent.click(screen.getByRole("button", { name: "Заменить" }));

    const dialog = screen.getByRole("dialog", { name: "Замена марки" });
    expect(within(dialog).queryByRole("textbox")).not.toBeInTheDocument();
    expect(within(dialog).queryByLabelText(/GUID/)).not.toBeInTheDocument();
    expect(preview).not.toHaveBeenCalled();

    fireEvent.click(within(dialog).getByRole("button", { name: exactLabel(`${name} · 11111111`) }));
    await waitFor(() => expect(preview).toHaveBeenCalledTimes(1));
    expect(preview.mock.calls[0][0].nomenclature_id).toBe(first);
    expect(preview.mock.calls[0][0].target_mark).toBe(name);

    const secondCard = within(dialog).getByRole("button", { name: exactLabel(`${name} · 22222222`) });
    await waitFor(() => expect(secondCard).toBeEnabled());
    fireEvent.click(secondCard);
    await waitFor(() => expect(preview).toHaveBeenCalledTimes(2));
    const write = await screen.findByRole("button", { name: "Записать замену" });
    await waitFor(() => expect(write).toBeEnabled());
    fireEvent.click(write);
    await waitFor(() => expect(commit).toHaveBeenCalledTimes(1));
    expect(commit.mock.calls[0][0].nomenclature_id).toBe(second);
    expect(commit.mock.calls[0][0].target_mark).toBe(name);
  });

  it("shows two different card names as buttons and does not append a GUID tail", async () => {
    const preview = vi.fn().mockResolvedValue({
      items: [
        {
          kp_id: 12,
          qty: 1,
          was: money(10),
          will: money(20),
          delivery_was: null,
          delivery_will: null,
          oneoff_notice: null,
        },
      ],
      skipped: [],
    });
    const commit = vi.fn().mockResolvedValue({ written: [], skipped: [], queue_notice: null });
    explainState.data = {
      ...EXPLAIN,
      product_kind: "bridge_pile",
      same_cards: [
        {
          mark: "С15-35T7",
          guid_count: 2,
          label: null,
          selectable: true,
          cards: [
            { name: "Сваи мостовые С15-35T7 завод", guid: "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaa11" },
            { name: "Сваи мостовые С15-35T7 особая", guid: "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbb22" },
          ],
        },
      ],
      slice: [],
    };
    render(
      <GuidExplainPanel
        source={{ ...SOURCE, scope: "economist", previewReplace: preview, commitReplace: commit }}
        productKind="bridge_pile"
        mark="C15-35T7"
      />,
    );
    open();
    fireEvent.click(screen.getByRole("button", { name: "Заменить" }));

    const dialog = screen.getByRole("dialog", { name: "Замена марки" });
    expect(within(dialog).queryByRole("textbox")).not.toBeInTheDocument();
    fireEvent.click(
      within(dialog).getByRole("button", { name: exactLabel("Сваи мостовые С15-35T7 особая") }),
    );
    expect(
      within(dialog).queryByRole("button", { name: /bbbbbb22|·/ }),
    ).not.toBeInTheDocument();

    await waitFor(() => expect(preview).toHaveBeenCalledTimes(1));
    expect(preview.mock.calls[0][0]).toMatchObject({
      target_mark: "С15-35T7",
      nomenclature_id: "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbb22",
    });
    expect(commit).not.toHaveBeenCalled();
  });
});

const exactLabel = (label: string) => (accessibleName: string) => accessibleName === label;

const money = (unit: number) => ({
  unit_price: unit,
  unit_discounted: unit,
  line_sum: unit * 2,
  total: unit * 2,
  subtotal: unit * 2,
  vat_amount: 0,
});
