import { useState } from "react";

import { canRemovePage, type PageStatus } from "@/features/commercial-offer/lib/multiPageSource";
import { SourceImageLightbox } from "@/features/commercial-offer/components/SourceImageLightbox";
import { Alert } from "@/shared/ui/Alert";

export type SourceImageGalleryPage = {
  id: string;
  name: string;
  previewUrl: string;
  status: PageStatus;
  errorMessage?: string;
};

type SourceImageGalleryProps = {
  pages: SourceImageGalleryPage[];
  activeId: string | null;
  onSelect: (id: string) => void;
  onRemove: (id: string) => void;
  /** When true, show remove+re-add hint if any page is in error. */
  showErrorHint?: boolean;
  /** R9: before OCR, thumbnail click opens large preview instead of only selecting. */
  enableLightbox?: boolean;
};

const STATUS_LABEL: Record<PageStatus, string> = {
  pending: "ожидает",
  running: "распознаётся",
  ready: "готово",
  error: "ошибка",
  confirmed: "подтверждено",
};

const STATUS_BADGE: Record<PageStatus, string | null> = {
  pending: "…",
  running: "⟳",
  ready: null,
  error: "!",
  confirmed: "✓",
};

export const SourceImageGallery = ({
  pages,
  activeId,
  onSelect,
  onRemove,
  showErrorHint = false,
  enableLightbox = false,
}: SourceImageGalleryProps) => {
  const [lightboxId, setLightboxId] = useState<string | null>(null);

  if (pages.length === 0) {
    return null;
  }

  const activePage = pages.find((page) => page.id === activeId) ?? null;
  const activeError =
    activePage?.status === "error" ? activePage.errorMessage ?? "Ошибка распознавания" : null;
  const hasErrorPage = pages.some((page) => page.status === "error");

  const handleSelect = (id: string) => {
    onSelect(id);
    if (enableLightbox) {
      setLightboxId(id);
    }
  };

  return (
    <div style={{ display: "grid", gap: "0.5rem" }}>
      <ul
        aria-label="Страницы источника"
        style={{
          display: "flex",
          gap: "0.75rem",
          listStyle: "none",
          margin: 0,
          padding: 0,
          flexWrap: "wrap",
        }}
      >
        {pages.map((page) => {
          const isActive = page.id === activeId;
          const removable = canRemovePage(page);
          const statusLabel = STATUS_LABEL[page.status];
          const badge = STATUS_BADGE[page.status];
          return (
            <li key={page.id} style={{ position: "relative" }}>
              <button
                type="button"
                onClick={() => handleSelect(page.id)}
                aria-label={`${page.name}, ${statusLabel}`}
                aria-pressed={isActive}
                style={{
                  display: "block",
                  width: 72,
                  height: 72,
                  padding: 0,
                  border: isActive
                    ? page.status === "error"
                      ? "2px solid #b42318"
                      : "2px solid #175cd3"
                    : page.status === "error"
                      ? "2px solid #fecdca"
                      : "2px solid #d0d5dd",
                  borderRadius: 10,
                  overflow: "hidden",
                  background: "#f2f4f7",
                  cursor: "pointer",
                }}
              >
                <img
                  src={page.previewUrl}
                  alt={page.name}
                  style={{ width: "100%", height: "100%", objectFit: "cover", display: "block" }}
                />
              </button>
              {badge && (
                <span
                  aria-label={`статус: ${statusLabel}`}
                  style={{
                    position: "absolute",
                    left: 4,
                    bottom: 4,
                    minWidth: 18,
                    height: 18,
                    borderRadius: 9,
                    background:
                      page.status === "error"
                        ? "#b42318"
                        : page.status === "confirmed"
                          ? "#067647"
                          : "#475467",
                    color: "#fff",
                    fontSize: 11,
                    lineHeight: "18px",
                    textAlign: "center",
                    padding: "0 4px",
                    pointerEvents: "none",
                  }}
                >
                  {badge}
                </span>
              )}
              {removable && (
                <button
                  type="button"
                  aria-label={`Удалить ${page.name}`}
                  onClick={(event) => {
                    event.stopPropagation();
                    onRemove(page.id);
                  }}
                  style={{
                    position: "absolute",
                    top: -6,
                    right: -6,
                    width: 22,
                    height: 22,
                    borderRadius: "50%",
                    border: "1px solid #d0d5dd",
                    background: "#fff",
                    color: "#b42318",
                    fontSize: 14,
                    lineHeight: "20px",
                    padding: 0,
                    cursor: "pointer",
                  }}
                >
                  ×
                </button>
              )}
            </li>
          );
        })}
      </ul>
      {activeError && <Alert tone="error">{activeError}</Alert>}
      {showErrorHint && hasErrorPage && (
        <Alert tone="warning">
          Удалите страницу с ошибкой и добавьте фото в конец
        </Alert>
      )}
      {enableLightbox && (
        <SourceImageLightbox
          pages={pages}
          openId={lightboxId}
          onClose={() => setLightboxId(null)}
          onNavigate={setLightboxId}
        />
      )}
    </div>
  );
};
