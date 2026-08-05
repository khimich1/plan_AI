export const FBS_GRADE_CODES = ["B7_5", "B20", "B22_5", "B25"] as const;

export type FbsGradeCode = (typeof FBS_GRADE_CODES)[number];

export const isFbsGradeCode = (value: string): value is FbsGradeCode =>
  (FBS_GRADE_CODES as readonly string[]).includes(value);

export const formatFbsGradeLabel = (code: string): string => {
  if (code === "B7_5") return "B7.5";
  if (code === "B20") return "B20";
  if (code === "B22_5") return "B22.5";
  if (code === "B25") return "B25";
  return code;
};
