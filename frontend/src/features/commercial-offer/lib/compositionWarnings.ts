const UNPARSED_COUNT_WARNING_RE = /^Не удалось распознать строк:\s*\d+$/;

export const filterCompositionWarnings = (warnings: string[]): string[] =>
  warnings.filter((warning) => !UNPARSED_COUNT_WARNING_RE.test(warning.trim()));
