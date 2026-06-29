import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { PRESET_TRACKS } from "./utils";

type Props = {
  tracksCount: number;
  canProceed: boolean;
  onTracksCountChange: (count: number) => void;
  onBack: () => void;
  onNext: () => void;
};

export const Step2TracksConfig = ({
  tracksCount,
  canProceed,
  onTracksCountChange,
  onBack,
  onNext,
}: Props) => (
  <div style={{ display: "grid", gap: "1rem" }}>
    <FieldWrapper
      label="Количество дорожек в день"
      hint="От 1 до 50. Максимум одновременно задействованных дорожек на одной дате."
    >
      <Input
        type="number"
        min={1}
        max={50}
        value={tracksCount}
        onChange={(e) =>
          onTracksCountChange(Math.max(1, Math.min(50, Number(e.target.value) || 1)))
        }
      />
    </FieldWrapper>

    <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
      {PRESET_TRACKS.map((preset) => (
        <Button
          key={preset}
          variant={tracksCount === preset ? "primary" : "secondary"}
          onClick={() => onTracksCountChange(preset)}
        >
          {preset}
        </Button>
      ))}
    </div>

    <div style={{ display: "flex", justifyContent: "space-between" }}>
      <Button variant="ghost" onClick={onBack}>
        ← Назад
      </Button>
      <Button onClick={onNext} disabled={!canProceed}>
        Далее →
      </Button>
    </div>
  </div>
);
