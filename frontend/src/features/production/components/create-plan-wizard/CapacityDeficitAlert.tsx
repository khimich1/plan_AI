import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import type {
  CapacityDeficit,
  CapacityOption,
} from "@/features/production/types/production";
import { formatRu } from "./utils";

export type CapacityDeficitAlertProps = {
  deficit: CapacityDeficit | null | undefined;
  applying?: boolean;
  onApplyOption: (option: CapacityOption) => void;
};

function optionLabel(option: CapacityOption): string {
  const dateRu = formatRu(option.date);
  if (option.action === "bump_fill") {
    return `+${option.add_tracks} дорожек на ${dateRu} (дозаполнить выбранный день)`;
  }
  return `+${option.add_tracks} дорожек на ${dateRu} (добавить день)`;
}

export const CapacityDeficitAlert = ({
  deficit,
  applying = false,
  onApplyOption,
}: CapacityDeficitAlertProps) => {
  if (!deficit || deficit.tracks_missing <= 0) {
    return null;
  }

  const { tracks_needed, tracks_available, tracks_missing, deficit_until, options } =
    deficit;

  return (
    <Alert tone="warning">
      <div style={{ display: "grid", gap: "0.75rem" }}>
        <div style={{ fontWeight: 600 }}>Дефицит ёмкости до {formatRu(deficit_until)}</div>
        <div style={{ fontSize: "0.9rem", display: "grid", gap: "0.25rem" }}>
          <div>Нужно дорожек: {tracks_needed}</div>
          <div>Доступно: {tracks_available}</div>
          <div>Не хватает: {tracks_missing}</div>
        </div>
        {options.length === 0 ? (
          <div style={{ fontSize: "0.9rem" }}>
            Нет свободных дней в пределах горизонта. Откройте календарь и выберите
            дни кистью.
          </div>
        ) : (
          <div style={{ display: "grid", gap: "0.5rem" }}>
            <div style={{ fontSize: "0.9rem", fontWeight: 600 }}>
              Варианты дозаполнения (выберите):
            </div>
            {options.map((option) => (
              <Button
                key={`${option.action}-${option.date}`}
                variant="secondary"
                onClick={() => {
                  if (option.action === "propose_day") {
                    const ok = window.confirm(
                      `Добавить день ${formatRu(option.date)} на ${option.add_tracks} дорожек в корзину?`,
                    );
                    if (!ok) return;
                  }
                  onApplyOption(option);
                }}
                disabled={applying || option.add_tracks <= 0}
              >
                {optionLabel(option)}
              </Button>
            ))}
          </div>
        )}
      </div>
    </Alert>
  );
};
