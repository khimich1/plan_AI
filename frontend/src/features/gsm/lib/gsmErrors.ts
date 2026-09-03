import { ApiError, getErrorMessage } from "@/shared/lib/apiError";

const EXACT_MESSAGES: Record<string, string> = {
  "tank_volume_liters must be > 0": "Объём бака должен быть больше 0.",
  "norm_summer must be > 0": "Летняя норма расхода должна быть больше 0.",
  "norm_winter must be > 0": "Зимняя норма расхода должна быть больше 0.",
  "tank and norms must be numbers > 0": "Бак и нормы должны быть числами больше 0.",
  "card_number is required": "Укажите номер топливной карты.",
  "card update conflict": "Конфликт при обновлении карты. Обновите список и повторите.",
  "hook_threshold_km must be > 0": "Порог крюка должен быть больше 0.",
  "station address already exists": "Станция с таким адресом уже есть.",
  "period_to must be >= period_from": "Дата «по» должна быть не раньше даты «с».",
  "vehicle has no primary_driver_id": "У машины не указан основной водитель.",
  "vehicle has no routes in library": "У машины нет маршрутов в библиотеке.",
  "fuel_start and odometer_start required when no confirmed waybill exists":
    "Укажите стартовый остаток бака и одометр — подтверждённого ПЛ ещё нет.",
  "last confirmed waybill missing fuel_end/odometer_end":
    "В последнем подтверждённом ПЛ нет остатка или одометра — задайте старт вручную.",
  "period has confirmed/exported waybills; pass force=true to overwrite":
    "В периоде есть подтверждённые ПЛ. Включите «Перезаписать» (force), чтобы продолжить.",
};

const CODE_MESSAGES: Record<string, string> = {
  gsm_validation: "Проверьте введённые данные.",
  gsm_card_duplicate: "Карта с таким номером уже существует.",
  gsm_station_duplicate: "Станция с таким адресом уже существует.",
  gsm_vehicle_not_found: "Машина не найдена.",
  gsm_driver_not_found: "Водитель не найден.",
  gsm_card_not_found: "Карта не найдена.",
  gsm_station_not_found: "Станция не найдена.",
  gsm_invalid_period: "Некорректный период генерации.",
  gsm_confirmed_conflict: "В периоде есть подтверждённые ПЛ — нужен флаг перезаписи.",
  gsm_unsolvable: "Генерация неразрешима: не хватает будних дней для баланса бака.",
  gsm_driver_required: "У машины не указан основной водитель.",
  gsm_routes_required: "У машины нет маршрутов в библиотеке.",
  gsm_start_required: "Укажите стартовый остаток бака и одометр — подтверждённого ПЛ ещё нет.",
  gsm_export_empty: "Нет путевых листов за выбранный период.",
  gsm_export_soffice_missing:
    "Не удалось экспортировать бланки: на сервере нет LibreOffice (soffice).",
  gsm_export_soffice_timeout: "Экспорт бланков превысил время ожидания LibreOffice.",
  gsm_export_soffice_failed: "LibreOffice не смог сформировать бланки путевых листов.",
  gsm_export_template_missing: "Не найден шаблон бланка путевого листа.",
  gsm_report_invalid_period: "Некорректный период отчёта.",
  gsm_report_no_data: "Нет подтверждённых путевых листов за выбранный период.",
  gsm_reset_no_anchors:
    "Не у всех активных машин есть imported-якорь — сброс невозможен.",
  gsm_reset_error: "Не удалось выполнить сброс ГСМ к якорям.",
};

export const formatGsmCodeMessage = (code: string | undefined, fallback: string): string => {
  if (code && CODE_MESSAGES[code]) {
    return CODE_MESSAGES[code];
  }
  const exact = EXACT_MESSAGES[fallback];
  if (exact) {
    return exact;
  }
  return fallback;
};

const translateDynamic = (message: string): string | null => {
  const cardDup = message.match(/^card_number «(.+)» already exists$/);
  if (cardDup) {
    return `Карта «${cardDup[1]}» уже существует.`;
  }
  const stationDup = message.match(/^station address «(.+)» already exists$/);
  if (stationDup) {
    return `Станция «${stationDup[1]}» уже существует.`;
  }
  const vehicleMissing = message.match(/^vehicle #(\d+) not found$/);
  if (vehicleMissing) {
    return `Машина №${vehicleMissing[1]} не найдена.`;
  }
  const driverMissing = message.match(/^driver #(\d+) not found$/);
  if (driverMissing) {
    return `Водитель №${driverMissing[1]} не найден.`;
  }
  const cardMissing = message.match(/^card #(\d+) not found$/);
  if (cardMissing) {
    return `Карта №${cardMissing[1]} не найдена.`;
  }
  if (message.includes("waybill is locked (confirmed/exported)")) {
    return "Путевой лист подтверждён или выгружен — редактирование запрещено.";
  }
  if (message.includes("cannot edit waybill: later confirmed/exported waybill exists")) {
    return "Нельзя править: после этого дня есть подтверждённые или выгруженные путевые.";
  }
  if (message.includes("season date must not be before last switch")) {
    return "Дата перевода сезона не может быть раньше предыдущего перевода.";
  }
  if (/insufficient_headroom|corridor_violation|cannot reach headroom|left corridor/.test(message)) {
    return "Генерация неразрешима: не хватает будних дней, чтобы удержать остаток бака.";
  }
  return null;
};

/** Human-readable GSM API / validation errors for UI alerts. */
export const formatGsmError = (error: unknown): string => {
  const raw = getErrorMessage(error);
  const exact = EXACT_MESSAGES[raw];
  if (exact) {
    return exact;
  }
  const dynamic = translateDynamic(raw);
  if (dynamic) {
    return dynamic;
  }
  if (error instanceof ApiError && error.code && CODE_MESSAGES[error.code]) {
    if (error.code.startsWith("gsm_export_")) {
      return CODE_MESSAGES[error.code];
    }
    // Prefer specific message when it is already Russian / useful; else code fallback.
    if (raw && raw !== "Запрос завершился ошибкой." && !/^[a-z_]+$/.test(raw)) {
      if (!EXACT_MESSAGES[raw] && !/must be|not found|already exists|is required/.test(raw)) {
        return raw;
      }
    }
    return CODE_MESSAGES[error.code];
  }
  return raw;
};
