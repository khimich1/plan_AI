import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import type { BasketDayKind } from "@/features/production/lib/basketDayKind";
import type { FillTargetItem } from "@/features/production/types/production";

const formatRu = (iso: string) => {
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y.slice(2)}`;
};

const clamp = (value: number, min: number, max: number): number => {
  if (Number.isNaN(value)) return min;
  return Math.max(min, Math.min(max, value));
};

export type FillBasketProps = {
  items: FillTargetItem[];
  basketKind: BasketDayKind | null;
  basketError?: string | null;
  /** Пресет кисти: число дорожек для новых дней. */
  brushTracks: number;
  onBrushTracksChange: (tracks: number) => void;
  /** Максимум для пресета (обычно max_per_day). */
  maxBrushTracks?: number;
  /** Правка N на чипе конкретного дня. */
  onChipTracksChange: (date: string, tracks: number) => void;
  /** Свободные слоты по дате — для clamp input на чипе. */
  freeSlotsByDate?: Record<string, number>;
  onRemove: (date: string) => void;
  onClear: () => void;
  onProceed: () => void;
};

/**
 * Sticky-плашка под календарём: пресет кисти N + чипы выбранных дней.
 * Всегда видна (даже при пустой корзине). CTA активен только при непустой корзине.
 */
export const FillBasket = ({
  items,
  basketKind,
  basketError,
  brushTracks,
  onBrushTracksChange,
  maxBrushTracks = 5,
  onChipTracksChange,
  freeSlotsByDate,
  onRemove,
  onClear,
  onProceed,
}: FillBasketProps) => {
  const isEmpty = items.length === 0;
  const totalTracks = items.reduce((acc, item) => acc + item.tracks, 0);
  const totalDays = items.length;
  const regionLabel =
    basketKind === "partial"
      ? "Корзина дозаполнения"
      : "Корзина планирования";
  const primaryLabel =
    basketKind === "partial"
      ? `Дозаполнить ${totalTracks} дор. на ${totalDays} дн. →`
      : "🚀 Начать планирование →";

  const maxPreset = Math.max(1, maxBrushTracks);

  return (
    <div className="fill-basket" role="region" aria-label={regionLabel}>
      {basketError ? (
        <div style={{ marginBottom: "0.5rem", width: "100%" }}>
          <Alert tone="warning">{basketError}</Alert>
        </div>
      ) : null}

      <div className="fill-basket__preset">
        <label className="fill-basket__preset-label">
          Дорожек:
          <input
            type="number"
            className="fill-basket__preset-input"
            min={1}
            max={maxPreset}
            value={brushTracks}
            aria-label="Пресет дорожек кисти"
            onChange={(e) =>
              onBrushTracksChange(clamp(Number(e.target.value), 1, maxPreset))
            }
          />
        </label>
        {isEmpty ? (
          <span className="fill-basket__hint">
            Клик — день в корзину · Shift+клик — диапазон · двойной клик / «i» — день
          </span>
        ) : null}
      </div>

      {!isEmpty ? (
        <div className="fill-basket__chips">
          {items.map((item) => {
            const maxChip = Math.max(
              1,
              freeSlotsByDate?.[item.date] ?? maxBrushTracks,
            );
            return (
              <span key={item.date} className="fill-basket__chip">
                <span>{formatRu(item.date)}</span>
                <input
                  type="number"
                  className="fill-basket__chip-input"
                  min={1}
                  max={maxChip}
                  value={item.tracks}
                  aria-label={`Дорожек на ${item.date}`}
                  onChange={(e) =>
                    onChipTracksChange(
                      item.date,
                      clamp(Number(e.target.value), 1, maxChip),
                    )
                  }
                />
                <span className="fill-basket__chip-unit">дор.</span>
                <button
                  type="button"
                  className="fill-basket__chip-remove"
                  aria-label={`Убрать ${item.date}`}
                  onClick={() => onRemove(item.date)}
                >
                  ✕
                </button>
              </span>
            );
          })}
        </div>
      ) : null}

      <div className="fill-basket__actions">
        {!isEmpty ? (
          <button type="button" className="fill-basket__clear" onClick={onClear}>
            Очистить
          </button>
        ) : null}
        <Button variant="primary" onClick={onProceed} disabled={isEmpty}>
          {isEmpty ? "Начать планирование →" : primaryLabel}
        </Button>
      </div>
    </div>
  );
};
