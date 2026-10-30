import { Button } from "@/shared/ui/Button";
import { FieldWrapper } from "@/shared/ui/Field";

export type TripBucket = "plate" | "pile" | "fbs_lm";

export type TripCountNotice = {
  previous: number;
  next: number;
};

const BUCKET_TITLE: Record<TripBucket, string> = {
  plate: "плит",
  pile: "свай",
  fbs_lm: "ФБС/ЛС/ЛМ",
};

const INPUT_LABEL: Record<TripBucket, string> = {
  plate: "Кол. рейсов плит",
  pile: "Кол. рейсов свай",
  fbs_lm: "Кол. рейсов ФБС/ЛС/ЛМ",
};

export function parseAcceptedTripCount(raw: string): number | null {
  const trimmed = raw.trim();
  if (!/^\d+$/.test(trimmed)) {
    return null;
  }
  const value = Number(trimmed);
  if (!Number.isInteger(value) || value < 0) {
    return null;
  }
  return value;
}

export function tripNoticeText(bucket: TripBucket, notice: TripCountNotice): string {
  return `Рейсы ${BUCKET_TITLE[bucket]} пересчитаны: было ${notice.previous}, стало ${notice.next}`;
}

export function TripCountNotices({
  notices,
  onDismiss,
  disabled = false,
}: {
  notices: Array<{ bucket: TripBucket; notice: TripCountNotice }>;
  onDismiss: (bucket: TripBucket) => void;
  disabled?: boolean;
}) {
  if (notices.length === 0) {
    return null;
  }
  return (
    <div style={{ display: "grid", gap: "0.45rem" }}>
      {notices.map(({ bucket, notice }) => (
        <div
          key={bucket}
          style={{
            display: "flex",
            justifyContent: "space-between",
            gap: "0.75rem",
            alignItems: "center",
            border: "1px solid #f2d39a",
            background: "#fff8eb",
            borderRadius: 10,
            padding: "0.55rem 0.75rem",
          }}
        >
          <span>{tripNoticeText(bucket, notice)}</span>
          <Button
            type="button"
            variant="ghost"
            disabled={disabled}
            aria-label={`Закрыть напоминание о рейсах ${BUCKET_TITLE[bucket]}`}
            onClick={() => onDismiss(bucket)}
          >
            Закрыть
          </Button>
        </div>
      ))}
    </div>
  );
}

export function TripCountField({
  bucket,
  value,
  formula,
  error,
  disabled = false,
  onChange,
  onOk,
}: {
  bucket: TripBucket;
  value: string;
  formula: number | null;
  error: string | null;
  disabled?: boolean;
  onChange: (value: string) => void;
  onOk: () => void;
}) {
  const parsed = parseAcceptedTripCount(value);
  const showHint = parsed !== null && formula !== null && parsed !== formula;
  const inputLabel = INPUT_LABEL[bucket];
  return (
    <div style={{ display: "grid", gap: "0.25rem", marginTop: "0.65rem" }}>
      <FieldWrapper label="Кол. рейсов" error={error}>
        <div style={{ position: "relative" }}>
          <input
            aria-label={inputLabel}
            value={value}
            onChange={(event) => onChange(event.target.value)}
            inputMode="numeric"
            disabled={disabled}
            style={{
              width: "100%",
              border: "1px solid #d0d5dd",
              borderRadius: 12,
              padding: "0.8rem 3.5rem 0.8rem 0.9rem",
              background: "#ffffff",
              boxSizing: "border-box",
            }}
          />
          <Button
            type="button"
            variant="secondary"
            aria-label={`OK ${inputLabel}`}
            onClick={onOk}
            disabled={disabled}
            style={{
              position: "absolute",
              right: "0.35rem",
              top: "50%",
              transform: "translateY(-50%)",
              borderRadius: 8,
              padding: "0.25rem 0.6rem",
              fontSize: "0.8rem",
            }}
          >
            OK
          </Button>
        </div>
      </FieldWrapper>
      {showHint ? (
        <span style={{ color: "#667085", fontSize: "0.85rem" }}>по весу: {formula}</span>
      ) : null}
    </div>
  );
}
