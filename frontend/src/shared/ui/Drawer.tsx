import { useEffect } from "react";
import type { PropsWithChildren, ReactNode } from "react";

type DrawerProps = PropsWithChildren<{
  open: boolean;
  onClose: () => void;
  title?: ReactNode;
  width?: number;
  footer?: ReactNode;
}>;

export const Drawer = ({
  open,
  onClose,
  title,
  width = 820,
  footer,
  children,
}: DrawerProps) => {
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
    const originalOverflow = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      window.removeEventListener("keydown", onKey);
      document.body.style.overflow = originalOverflow;
    };
  }, [open, onClose]);

  if (!open) {
    return null;
  }

  return (
    <div
      className="app-drawer"
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
