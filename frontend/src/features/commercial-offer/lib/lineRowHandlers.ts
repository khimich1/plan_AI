export type LineSavePayload = {
  qty?: number;
  sourceText?: string;
};

export type LineUndoToastState = {
  message: string;
  onUndo: () => void;
};

export type LineRowErrorState = {
  lineId: string;
  message: string;
};

export type LineRowHandlers = {
  onSaveLine: (lineId: string, payload: LineSavePayload) => void | Promise<void>;
  onDeleteLine: (lineId: string) => void | Promise<void>;
  undoToast: LineUndoToastState | null;
  rowError?: LineRowErrorState | null;
};

export const LINE_UNDO_TOAST_MS = 8000;
