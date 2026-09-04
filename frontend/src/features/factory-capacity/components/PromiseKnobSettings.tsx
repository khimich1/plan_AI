import { useEffect, useState, type CSSProperties } from "react";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Alert } from "@/shared/ui/Alert";
import { getErrorMessage } from "@/shared/lib/apiError";
import { useUpdatePromiseKnobMutation } from "@/features/factory-capacity/api/promiseQuote";

type Props = {
  currentKnob: number;
  min?: number;
  max?: number;
};

const wrapStyle: CSSProperties = {
  display: "grid",
  gap: "0.75rem",
  marginTop: "1rem",
  padding: "0.85rem 1rem",
  border: "1px solid #e4e7ec",
  borderRadius: 12,
  background: "#fafafa",
};

const hintStyle: CSSProperties = {
  color: "#667085",
  fontSize: "0.85rem",
  margin: 0,
};

export const PromiseKnobSettings = ({ currentKnob, min = 1, max = 5 }: Props) => {
  const [open, setOpen] = useState(false);
  const [draft, setDraft] = useState(String(currentKnob));
  const mutation = useUpdatePromiseKnobMutation();

  useEffect(() => {
    setDraft(String(currentKnob));
  }, [currentKnob]);

  const parsed = Number(draft);
  const valid = Number.isInteger(parsed) && parsed >= min && parsed <= max;
  const changed = valid && parsed !== currentKnob;

  const onConfirm = async () => {
    if (!changed || mutation.isPending) {
      return;
    }
    try {
      await mutation.mutateAsync(parsed);
      setOpen(false);
    } catch {
      // ошибка через mutation.error
    }
  };

  return (
    <section data-testid="promise-knob-settings" style={wrapStyle}>
      {!open ? (
        <Button
          type="button"
          variant="ghost"
          onClick={() => {
            mutation.reset();
            setDraft(String(currentKnob));
            setOpen(true);
          }}
        >
          Настроить ручку ({currentKnob} дор./день)
        </Button>
      ) : (
        <form
          onSubmit={(event) => {
            event.preventDefault();
            void onConfirm();
          }}
          style={{ display: "grid", gap: "0.75rem" }}
        >
          <FieldWrapper label="Дорожек в день">
            <Input
              type="number"
              min={min}
              max={max}
              step={1}
              value={draft}
              onChange={(event) => setDraft(event.target.value)}
              aria-label="Дорожек в день"
            />
          </FieldWrapper>
          <p style={hintStyle}>Влияет только на новые расчёты</p>
          {mutation.isError ? <Alert tone="error">{getErrorMessage(mutation.error)}</Alert> : null}
          <div style={{ display: "flex", gap: "0.5rem", justifyContent: "flex-end" }}>
            <Button
              type="button"
              variant="ghost"
              onClick={() => {
                setOpen(false);
                setDraft(String(currentKnob));
                mutation.reset();
              }}
              disabled={mutation.isPending}
            >
              Отмена
            </Button>
            <Button type="submit" disabled={!changed || mutation.isPending}>
              {mutation.isPending ? "Сохраняю…" : "Подтвердить"}
            </Button>
          </div>
        </form>
      )}
    </section>
  );
};
