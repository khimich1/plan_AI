import { useEffect } from "react";
import type { PropsWithChildren, ReactNode } from "react";

type DrawerProps = PropsWithChildren<{
  open: boolean;
  onClose: () => void;
  title?: ReactNode;
  width?: number;
  footer?: ReactNode;
  /** Viewport edge the panel attaches to. Default right (existing drawers). */
  side?: "left" | "right";
}>;

export const Drawer = ({
  open,
  onClose,
  title,
  width = 820,
  footer,
  side = "right",
  children,
}: DrawerProps) => {
  useEffect(() => {
    if (!open) {
      return;
    }
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.stopImmediatePropagation();
        onClose();
      }
    };
    // Capture: close this drawer before any underlying Modal Esc handler.
    window.addEventListener("keydown", onKey, true);
    const originalOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      window.removeEventListener("keydown", onKey, true);
      document.body.style.overflow = originalOverflow;
    };
  }, [open, onClose]);

  if (!open) {
    return null;
  }

  const sideClass = side === "left" ? "app-drawer app-drawer--left" : "app-drawer";

  return (
    <div
      className={sideClass}
      role="dialog"
      aria-modal="true"
      onClick={onClose}
    >
      <div
        className="app-drawer__panel"
        style={{ width, maxWidth: "100vw" }}
        onClick={(event) => event.stopPropagation()}
      >
        {(title || true) && (
          <header className="app-drawer__header">
            <div className="app-drawer__title">{title}</div>
            <button
              type="button"
              onClick={onClose}
              aria-label="Закрыть"
              className="app-drawer__close"
            >
              ×
            </button>
          </header>
        )}
        <div className="app-drawer__body">{children}</div>
        {footer && <footer className="app-drawer__footer">{footer}</footer>}
      </div>
    </div>
  );
};
