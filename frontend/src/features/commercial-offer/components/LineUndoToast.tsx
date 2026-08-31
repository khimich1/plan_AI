import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";

type LineUndoToastProps = {
  message: string;
  onUndo: () => void;
};

export const LineUndoToast = ({ message, onUndo }: LineUndoToastProps) => (
  <Alert tone="info">
    <div
      style={{
        display: "flex",
        flexWrap: "wrap",
        alignItems: "center",
        justifyContent: "space-between",
        gap: "0.75rem",
      }}
    >
      <span>{message}</span>
      <Button type="button" variant="ghost" onClick={onUndo}>
        Отменить
      </Button>
    </div>
  </Alert>
);
