export const PILE_GRADE_CODES = ["B15", "B20", "B22_5", "B25", "B30_granite"] as const;

export type PileGradeCode = (typeof PILE_GRADE_CODES)[number];

export const formatPileGradeLabel = (code: string): string => code.replace("_", ".");

export const isPileGradeCode = (value: string): value is PileGradeCode =>
  (PILE_GRADE_CODES as readonly string[]).includes(value);
