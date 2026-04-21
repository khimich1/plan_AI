export class ApiError extends Error {
  status: number;
  detail: string;

  constructor(message: string, status = 500, detail = message) {
    super(message);
    this.name = "ApiError";
    this.status = status;
    this.detail = detail;
  }
}

export const getErrorMessage = (error: unknown): string => {
  if (error instanceof ApiError) {
    return error.detail || error.message;
  }
  if (error instanceof Error) {
    return error.message;
  }
  return "Неизвестная ошибка. Попробуйте ещё раз.";
};
