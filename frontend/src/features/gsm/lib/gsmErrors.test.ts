import { describe, expect, it } from "vitest";
import { ApiError } from "@/shared/lib/apiError";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";

describe("formatGsmError", () => {
  it("translates known English validation messages", () => {
    expect(formatGsmError(new ApiError("tank", 422, "tank_volume_liters must be > 0"))).toBe(
      "Объём бака должен быть больше 0.",
    );
    expect(formatGsmError(new ApiError("norm", 422, "norm_summer must be > 0"))).toBe(
      "Летняя норма расхода должна быть больше 0.",
    );
  });

  it("translates duplicate card message with number", () => {
    expect(
      formatGsmError(new ApiError("dup", 422, "card_number «7001» already exists", "gsm_card_duplicate")),
    ).toBe("Карта «7001» уже существует.");
  });

  it("falls back to code message when detail is technical", () => {
    expect(
      formatGsmError(new ApiError("x", 404, "vehicle #9 not found", "gsm_vehicle_not_found")),
    ).toBe("Машина №9 не найдена.");
  });

  it("translates locked waybill conflict", () => {
    expect(
      formatGsmError(new ApiError("locked", 409, "waybill is locked (confirmed/exported)")),
    ).toBe("Путевой лист подтверждён или выгружен — редактирование запрещено.");
  });

  it("translates later confirmed/exported waybill conflict with extra detail", () => {
    expect(
      formatGsmError(
        new ApiError(
          "locked",
          409,
          "cannot edit waybill: later confirmed/exported waybill exists (2026-08-20)",
        ),
      ),
    ).toBe("Нельзя править: после этого дня есть подтверждённые или выгруженные путевые.");
  });

  it("translates season switch validation error", () => {
    expect(
      formatGsmError(new ApiError("season", 422, "season date must not be before last switch")),
    ).toBe("Дата перевода сезона не может быть раньше предыдущего перевода.");
  });

  it("translates LibreOffice export failures with an admin hint", () => {
    expect(
      formatGsmError(
        new ApiError(
          "soffice",
          500,
          "LibreOffice (soffice) is not installed or not on PATH",
          "gsm_export_soffice_missing",
        ),
      ),
    ).toBe("Не удалось экспортировать бланки: на сервере нет LibreOffice (soffice).");
  });
});
