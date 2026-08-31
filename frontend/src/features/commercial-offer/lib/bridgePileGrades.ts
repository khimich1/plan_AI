export const BRIDGE_PILE_GRADE_CODES = ["B25", "B30"] as const;

export type BridgePileGradeCode = (typeof BRIDGE_PILE_GRADE_CODES)[number];

export const isBridgePileGradeCode = (value: string): value is BridgePileGradeCode =>
  (BRIDGE_PILE_GRADE_CODES as readonly string[]).includes(value);

export const formatBridgePileGradeLabel = (code: string): string => {
  if (code === "B25") return "B25";
  if (code === "B30") return "B30";
  return code;
};
