import { useEffect, useState, type WheelEvent } from "react";

import {
  clampQueueIndex,
  nextIndex,
  prevIndex,
  type SourceImageQueueItem,
} from "@/features/commercial-offer/lib/sourceImageQueue";
import { Button } from "@/shared/ui/Button";
import { Drawer } from "@/shared/ui/Drawer";

type SourceImageQueueDrawerProps = {
  open: boolean;
  onClose: () => void;
  items: SourceImageQueueItem[];
};

const IMAGE_ZOOM_MIN = 0.5;
const IMAGE_ZOOM_MAX = 3;
const IMAGE_ZOOM_STEP = 0.25;
const IMAGE_ZOOM_FIT = 1;

const clampImageZoom = (value: number) =>
  Math.min(IMAGE_ZOOM_MAX, Math.max(IMAGE_ZOOM_MIN, value));

const formatImageZoom = (zoom: number) => `${Math.round(zoom * 100)}%`;

const compactButtonStyle = { minWidth: 32, padding: "0.25rem 0.5rem" } as const;

export const SourceImageQueueDrawer = ({
  open,
  onClose,
  items,
}: SourceImageQueueDrawerProps) => {
  const [index, setIndex] = useState(0);
  const [imageZoom, setImageZoom] = useState(IMAGE_ZOOM_FIT);

  useEffect(() => {
    if (!open) {
      setIndex(0);
      setImageZoom(IMAGE_ZOOM_FIT);
      return;
    }
    setIndex((current) => clampQueueIndex(current, items.length));
  }, [open, items.length]);

  const stepIndex = (compute: (current: number) => number) => {
    setIndex(compute);
    setImageZoom(IMAGE_ZOOM_FIT);
  };

  if (!open || items.length === 0) {
    return null;
  }

  const safeIndex = clampQueueIndex(index, items.length);
  const current = items[safeIndex]!;
  const showPager = items.length > 1;

  const bumpZoom = (direction: 1 | -1) => {
    setImageZoom((currentZoom) =>
      clampImageZoom(Number((currentZoom + direction * IMAGE_ZOOM_STEP).toFixed(2))),
    );
  };

  const handleImageWheel = (event: WheelEvent<HTMLDivElement>) => {
    if (!event.ctrlKey) {
      return;
    }
    event.preventDefault();
    bumpZoom(event.deltaY < 0 ? 1 : -1);
  };

  return (
    <Drawer open={open} onClose={onClose} title="Исходные фото" side="left" width={560}>
      <div style={{ display: "grid", gap: "0.75rem" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "0.75rem" }}>
          <div style={{ fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis" }}>{current.name}</div>
          <div style={{ color: "#667085", whiteSpace: "nowrap" }}>
            {safeIndex + 1} / {items.length}
          </div>
        </div>

        <div
          onWheel={handleImageWheel}
          style={{
            overflow: "auto",
            width: "100%",
            maxHeight: "70vh",
            minHeight: 280,
            borderRadius: 12,
            border: "1px solid #e4e7ec",
            background: "#f8fafc",
          }}
        >
          <img
            src={current.url}
            alt={current.name}
            style={{
              display: "block",
              width: `${imageZoom * 100}%`,
              height: "auto",
              maxWidth: "none",
            }}
          />
        </div>

        <div style={{ display: "flex", flexWrap: "wrap", gap: "0.5rem", alignItems: "center" }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: "0.25rem",
              border: "1px solid #d0d5dd",
              borderRadius: 8,
              padding: "0.15rem",
              background: "#ffffff",
            }}
          >
            <Button
              type="button"
              variant="ghost"
              aria-label="Уменьшить"
              title="Уменьшить"
              onClick={() => bumpZoom(-1)}
              disabled={imageZoom <= IMAGE_ZOOM_MIN}
              style={compactButtonStyle}
            >
              −
            </Button>
            <span style={{ minWidth: 48, textAlign: "center", fontSize: "0.85rem", color: "#475467" }}>
              {formatImageZoom(imageZoom)}
            </span>
            <Button
              type="button"
              variant="ghost"
              aria-label="Увеличить"
              title="Увеличить"
              onClick={() => bumpZoom(1)}
              disabled={imageZoom >= IMAGE_ZOOM_MAX}
              style={compactButtonStyle}
            >
              +
            </Button>
            <Button
              type="button"
              variant="ghost"
              aria-label="По ширине"
              title="По ширине окна"
              onClick={() => setImageZoom(IMAGE_ZOOM_FIT)}
              style={{ padding: "0.25rem 0.5rem", fontSize: "0.85rem" }}
            >
              По ширине
            </Button>
          </div>

          {showPager && (
            <>
              <Button
                type="button"
                variant="secondary"
                aria-label="Предыдущее фото"
                onClick={() => stepIndex((currentIndex) => prevIndex(currentIndex, items.length))}
              >
                ←
              </Button>
              <Button
                type="button"
                variant="secondary"
                aria-label="Следующее фото"
                onClick={() => stepIndex((currentIndex) => nextIndex(currentIndex, items.length))}
              >
                →
              </Button>
            </>
          )}
          <a
            href={current.url}
            target="_blank"
            rel="noreferrer"
            style={{ color: "#175cd3", marginLeft: "auto" }}
          >
            Открыть в новой вкладке
          </a>
        </div>
        <div style={{ fontSize: "0.8rem", color: "#667085" }}>Ctrl + колёсико мыши — масштаб</div>
      </div>
    </Drawer>
  );
};
