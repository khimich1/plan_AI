export type DbStatsResponse = {
  kp_total: number;
  kp_in_work: number;
  kp_completed: number;
  plates_in_work: number;
  plates_completed: number;
  plate_rests: number;
  plans_count: number;
  current_plan_present: boolean;
};

export type DbResetReport = {
  sqlite: Record<string, number>;
  plans: Record<string, number>;
  calendar_reset: boolean;
};

export type RecoverPlatesResponse = {
  recovered_records: number;
};

export type ResetVariant =
  | "full"
  | "kp-only"
  | "plans-only"
  | "calendar-only";
