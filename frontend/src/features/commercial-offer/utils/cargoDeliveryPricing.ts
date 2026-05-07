/** Максимальная масса груза на один рейс (кг), совпадает с backend `CARGO_DELIVERY_TRUCK_CAPACITY_KG`. */
export const CARGO_DELIVERY_TRUCK_CAPACITY_KG = 18600;

export function cargoDeliveryTripsCount(cargoWeightKg: number): number {
  const w = Math.max(0, cargoWeightKg);
  if (w <= 0) {
    return 0;
  }
  return Math.ceil(w / CARGO_DELIVERY_TRUCK_CAPACITY_KG);
}

export function deliveryServiceTotalRub(tripCostRub: number, cargoWeightKg: number): number {
  const trip = Math.max(0, tripCostRub);
  const n = cargoDeliveryTripsCount(cargoWeightKg);
  return Math.round(trip * n * 100) / 100;
}
