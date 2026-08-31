import { describe, expect, it } from "vitest";
import { formatLineSourceText } from "@/features/commercial-offer/lib/formatLineSourceText";

describe("formatLineSourceText", () => {
  it("formats a plate row as mark plus qty", () => {
    expect(
      formatLineSourceText({
        product_type: "plates",
        name: "Плиты ПБ 78-12-8п",
        qty: 2,
      }),
    ).toBe("ПБ 78-12-8п 2");
  });

  it("formats a pile row with grade", () => {
    expect(
      formatLineSourceText({
        product_type: "piles",
        mark: "С120.35-12",
        concrete_grade: "B25",
        qty: 5,
      }),
    ).toBe("С120.35-12 B25 5");
  });

  it("formats a step row without grade", () => {
    expect(
      formatLineSourceText({
        product_type: "steps",
        mark: "ЛС 12",
        qty: 4,
      }),
    ).toBe("ЛС 12 4");
  });
});
