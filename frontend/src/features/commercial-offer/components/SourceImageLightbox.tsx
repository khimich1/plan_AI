import { useEffect } from "react";

export type SourceImageLightboxPage = {
  id: string;
  name: string;
  previewUrl: string;
};

type SourceImageLightboxProps = {
  pages: SourceImageLightboxPage[];
  openId: string | null;
  onClose: () => void;
  onNavigate: (id: string) => void;
};

export const SourceImageLightbox = ({
  pages,
  openId,
  onClose,
  onNavigate,
}: SourceImageLightboxProps) => {
  const index = openId ? pages.findIndex((page) => page.id === openId) : -1;
  const page = index >= 0 ? pages[index] : null;
  const prevId = index > 0 ? pages[index - 1]?.id : null;
  const nextId = index >= 0 && index < pages.length - 1 ? pages[index + 1]?.id : null;

  useEffect(() => {
    if (!openId) {
      return;
    }
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        onClose();
        return;
      }
      if (event.key === "ArrowLeft" && prevId) {
        onNavigate(prevId);
        return;
      }
      if (event.key === "ArrowRight" && nextId) {
        onNavigate(nextId);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [openId, onClose, onNavigate, prevId, nextId]);

  if (!page) {
    return null;
  }

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label={page.name}
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(15, 23, 42, 0.72)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 1100,
        padding: "1rem",
      }}
    >
      <div
        onClick={(event) => event.stopPropagation()}
        style={{
          position: "relative",
          maxWidth: "min(960px, 100%)",
          maxHeight: "90vh",
          display: "grid",
          gap: "0.75rem",
          justifyItems: "center",
        }}
      >
        <div
          style={{
            display: "flex",
            width: "100%",
            justifyContent: "space-between",
            alignItems: "center",
            gap: "0.75rem",
            color: "#f8fafc",
          }}
        >
          <span style={{ fontSize: "0.95rem" }}>
            {page.name}
            {pages.length > 1 ? ` · ${index + 1}/${pages.length}` : ""}
          </span>
          <button
            type="button"
            onClick={onClose}
            aria-label="Закрыть"
            style={{
              background: "transparent",
              border: "none",
              fontSize: "1.5rem",
              lineHeight: 1,
              cursor: "pointer",
              color: "#f8fafc",
              padding: "0.25rem 0.5rem",
            }}
          >
            ×
          </button>
        </div>

        <img
          src={page.previewUrl}
          alt={page.name}
          style={{
            maxWidth: "100%",
            maxHeight: "75vh",
            objectFit: "contain",
            borderRadius: 8,
            background: "#0f172a",
          }}
        />

        {pages.length > 1 && (
          <div style={{ display: "flex", gap: "0.75rem" }}>
            <button
              type="button"
              onClick={() => prevId && onNavigate(prevId)}
              disabled={!prevId}
              aria-label="Предыдущая страница"
              style={{
                border: "1px solid #94a3b8",
                background: prevId ? "#1e293b" : "#334155",
                color: "#f8fafc",
                borderRadius: 8,
                padding: "0.4rem 0.9rem",
                cursor: prevId ? "pointer" : "not-allowed",
                opacity: prevId ? 1 : 0.5,
              }}
            >
              ←
            </button>
            <button
              type="button"
              onClick={() => nextId && onNavigate(nextId)}
              disabled={!nextId}
              aria-label="Следующая страница"
              style={{
                border: "1px solid #94a3b8",
                background: nextId ? "#1e293b" : "#334155",
                color: "#f8fafc",
                borderRadius: 8,
                padding: "0.4rem 0.9rem",
                cursor: nextId ? "pointer" : "not-allowed",
                opacity: nextId ? 1 : 0.5,
              }}
            >
              →
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
