import { Button } from "@/shared/ui/Button";
import type { FillTargetItem } from "@/features/production/types/production";

const formatRu = (iso: string) => {
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y.slice(2)}`;
};

export type FillBasketProps = {
  items: FillTargetItem[];
  onRemove: (date: string) => void;
  onClear: () => void;
  onProceed: () => void;
};

/**
 * Sticky-плашка под календарём с чипами выбранных дней.
 *
 * Редактирование количества дорожек делается не здесь, а через
 * `DayDrawer`: пользователь кликает день → меняет N → жмёт «Заменить».
 * Это снижает количество мест, где можно случайно расхождение получить.
 */
export const FillBasket = ({ items, onRemove, onClear, onProceed }: FillBasketProps) => {
  if (items.length === 0) return null;

  const totalTracks = items.reduce((acc, item) => acc + item.tracks, 0);
  const totalDays = items.length;

  return (
    <div className="fill-basket" role="region" aria-label="Корзина дозаполнения">
      <div className="fill-basket__chips">
        {items.map((item) => (
          <span key={item.date} className="fill-basket__chip">
            <span>
              {formatRu(item.date)} · {item.tracks} дор.
            </span>
            <button
              type="button"
              className="fill-basket__chip-remove"
              aria-label={`Убрать ${item.date}`}
              onClick={() => onRemove(item.date)}
            >
              ✕
            </button>
          </span>
        ))}
      </div>
      <div className="fill-basket__actions">
        <button type="button" className="fill-basket__clear" onClick={onClear}>
          Очистить
        </button>
        <Button variant="primary" onClick={onProceed}>
          Дозаполнить {totalTracks} дор. на {totalDays} дн. →
        </Button>
      </div>
    </div>
  );
};
