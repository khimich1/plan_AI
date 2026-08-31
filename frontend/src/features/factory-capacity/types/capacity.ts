/** Types for factory capacity snapshot (GET .../capacity-snapshot). */

export type CapacityStatus = "green" | "yellow" | "red";

export type CapacityDayInfo = {
  occupied: number;
  max: number;
};

export type CapacitySnapshot = {
  start_date: string;
  target_date: string;
  tracks_needed: number;
  tracks_free_in_window: number;
  delta: number;
  status: CapacityStatus;
  hint: string | null;
  days_info: Record<string, CapacityDayInfo>;
  holidays: string[];
  extra_workdays: string[];
  calendar_from_month: string;
  calendar_to_month: string;
};
