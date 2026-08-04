import type { SgpPlateItem } from "@/features/production/api/sgpApi";

type UnlinkPlate = Pick<SgpPlateItem, "kp_id" | "plate_name" | "qty">;

/** qty=1: confirm-only flow without a quantity input. */
export const isSingleQtyUnlink = (qty: number): boolean => qty === 1;

export const getUnlinkPrompt = (plate: UnlinkPlate): string => {
  if (isSingleQtyUnlink(plate.qty)) {
    return `Отвязать 1 шт от КП #${plate.kp_id}?`;
  }
  return `Отвязать от КП #${plate.kp_id}: ${plate.plate_name}`;
};

export const getUnlinkConfirmButtonLabel = (qty: number): string =>
  isSingleQtyUnlink(qty) ? "Да, отвязать" : "Подтвердить";

/** Quantity sent to API: always 1 for single-qty rows. */
export const resolveUnlinkSubmitQty = (plate: Pick<SgpPlateItem, "qty">, unlinkQty: number): number =>
  isSingleQtyUnlink(plate.qty) ? 1 : unlinkQty;
