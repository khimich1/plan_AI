import { Button } from "@/shared/ui/Button";
import type { PageSource } from "@/features/commercial-offer/lib/multiPageSource";
import { SourceImageGallery } from "@/features/commercial-offer/components/SourceImageGallery";

type PageReviewNavProps = {
  pages: PageSource[];
  activeId: string | null;
  progressLabel: string | null;
  onSelect: (id: string) => void;
  onRemove?: (id: string) => void;
  onPrev: () => void;
  onNext: () => void;
};

export const PageReviewNav = ({
  pages,
  activeId,
  progressLabel,
  onSelect,
  onRemove,
  onPrev,
  onNext,
}: PageReviewNavProps) => {
  if (pages.length <= 1 && !progressLabel) {
    return null;
  }

  const navigable = pages.filter((page) => page.status === "ready" || page.status === "confirmed");
  const activeIndex = navigable.findIndex((page) => page.id === activeId);
  const canPrev = activeIndex > 0;
  const canNext = activeIndex >= 0 && activeIndex < navigable.length - 1;

  return (
    <div style={{ display: "grid", gap: "0.75rem" }}>
      <div style={{ display: "flex", gap: "0.75rem", flexWrap: "wrap", alignItems: "center" }}>
        {progressLabel && <span style={{ fontSize: "0.9rem", color: "#475467" }}>{progressLabel}</span>}
        {pages.length > 1 && (
          <>
            <Button type="button" variant="ghost" onClick={onPrev} disabled={!canPrev} title="Предыдущая страница">
              ←
            </Button>
            <Button type="button" variant="ghost" onClick={onNext} disabled={!canNext} title="Следующая страница">
              →
            </Button>
          </>
        )}
      </div>
      {pages.length > 1 && (
        <SourceImageGallery
          pages={pages}
          activeId={activeId}
          onSelect={(id) => {
            const page = pages.find((item) => item.id === id);
            if (page && (page.status === "ready" || page.status === "confirmed" || page.status === "error")) {
              onSelect(id);
            }
          }}
          onRemove={onRemove ?? (() => undefined)}
          showErrorHint
        />
      )}
    </div>
  );
};
