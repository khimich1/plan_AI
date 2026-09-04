import { useEffect, useState } from "react";

import { SourceImageQueueDrawer } from "@/features/commercial-offer/components/SourceImageQueueDrawer";
import type { SourceImageQueueItem } from "@/features/commercial-offer/lib/sourceImageQueue";
import { Button } from "@/shared/ui/Button";

type SourceImageQueueControlsProps = {
  items?: SourceImageQueueItem[];
};

/** CTA «Исходные фото (N)» + left Drawer; hidden when queue is empty. */
export const SourceImageQueueControls = ({ items = [] }: SourceImageQueueControlsProps) => {
  const [open, setOpen] = useState(false);

  useEffect(() => {
    if (items.length === 0) {
      setOpen(false);
    }
  }, [items.length]);

  return (
    <>
      {items.length > 0 && (
        <div>
          <Button
            type="button"
            variant="secondary"
            aria-label={`Исходные фото (${items.length})`}
            onClick={() => setOpen(true)}
          >
            Исходные фото ({items.length})
          </Button>
        </div>
      )}
      <SourceImageQueueDrawer open={open} onClose={() => setOpen(false)} items={items} />
    </>
  );
};
