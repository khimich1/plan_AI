import { describe, expect, it } from "vitest";
import {
  getUnlinkConfirmButtonLabel,
  getUnlinkPrompt,
  isSingleQtyUnlink,
  resolveUnlinkSubmitQty,
} from "@/features/production/lib/sgpWarehouseUnlink";

describe("sgpWarehouseUnlink", () => {
  const linkedPlate = {
    kp_id: 1,
    plate_name: "Плиты ПБ 45-12-6п",
    qty: 1,
  };

  const multiQtyPlate = {
    kp_id: 2,
    plate_name: "Плиты ПБ 74,1-12-8п",
    qty: 5,
  };

  it("treats qty=1 as single-qty unlink", () => {
    expect(isSingleQtyUnlink(1)).toBe(true);
    expect(isSingleQtyUnlink(2)).toBe(false);
  });

  it("uses short confirm prompt for single-qty unlink", () => {
    expect(getUnlinkPrompt(linkedPlate)).toBe("Отвязать 1 шт от КП #1?");
  });

  it("uses detailed prompt with plate name when qty>1", () => {
    expect(getUnlinkPrompt(multiQtyPlate)).toBe(
      "Отвязать от КП #2: Плиты ПБ 74,1-12-8п",
    );
  });

  it("uses fast confirm label for single-qty unlink", () => {
    expect(getUnlinkConfirmButtonLabel(1)).toBe("Да, отвязать");
    expect(getUnlinkConfirmButtonLabel(3)).toBe("Подтвердить");
  });

  it("submits qty=1 for single-qty rows regardless of draft input", () => {
    expect(resolveUnlinkSubmitQty(linkedPlate, 99)).toBe(1);
    expect(resolveUnlinkSubmitQty(multiQtyPlate, 3)).toBe(3);
  });
});
