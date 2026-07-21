import { useEffect } from "react";
import type { PropsWithChildren, ReactNode } from "react";

type ModalProps = PropsWithChildren<{
  open: boolean;
  onClose: () => void;
  title?: ReactNode;
  maxWidth?: number;
}>;

export const Modal = ({ open, onClose, title, maxWidth = 560, children }: ModalProps) => {
  useEffect(() => {
    if (!open) {
      return;
    }
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        onClose();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  if (!open) {
    return null;
  }

  return (
    <div
      role="dialog"
      aria-modal="true"
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(15, 23, 42, 0.45)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 1000,
        padding: "1rem",
      }}
    >
      <div
        onClick={(event) => event.stopPropagation()}
        style={{
          background: "#ffffff",
          borderRadius: 20,
          boxShadow: "0 20px 60px rgba(15, 23, 42, 0.2)",
          padding: "1.5rem",
          width: "100%",
          maxWidth,
          maxHeight: "90vh",
          overflow: "auto",
        }}
      >
        {title && (
          <header style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
            <h2 style={{ margin: 0, fontSize: "1.25rem" }}>{title}</h2>
            <button
              type="button"
              onClick={onClose}
              aria-label="Закрыть"
              style={{
                background: "transparent",
                border: "none",
                fontSize: "1.25rem",
                cursor: "pointer",
                color: "#667085",
              }}
            >
              ×
            </button>
          </header>
        )}
        {children}
      </div>
    </div>
  );
};
