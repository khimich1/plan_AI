import { describe, expect, it } from "vitest";
import { filterCompositionWarnings } from "@/features/commercial-offer/lib/compositionWarnings";

describe("filterCompositionWarnings", () => {
  it("drops the unparsed-count banner and keeps other warnings", () => {
    const warnings = [
      "Не удалось распознать строк: 2",
      "Строки формата «длина×ширина×толщина» (мм), например «3880x1200x220»: нагрузка принята 8п по умолчанию. Проверьте нагрузку перед отправкой КП.",
    ];

    expect(filterCompositionWarnings(warnings)).toEqual([warnings[1]]);
  });
});
