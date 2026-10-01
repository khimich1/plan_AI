import { useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import type {
  SpecificationChoice,
  SpecificationChoicePayload,
  SpecificationDeliveryKind,
  SpecificationPaymentKind,
  SpecificationTermKind,
  SpecificationView,
} from "@/features/commercial-archive/types/archive";

const STALE_SPECIFICATION =
  "Условия оплаты в спецификации устарели: откройте спецификацию и сохраните их снова";

const PAYMENT_OPTIONS: { kind: SpecificationPaymentKind; label: string }[] = [
  { kind: "prepay_100", label: "100% предоплата" },
  { kind: "split_50_50", label: "50% до даты и 50% перед отгрузкой" },
  { kind: "split_share", label: "50% + доля + остаток после УПД" },
  { kind: "deferral", label: "Отсрочка" },
  { kind: "custom", label: "Свой текст" },
];

const TERM_OPTIONS: { kind: SpecificationTermKind; label: string }[] = [
  { kind: "by_date", label: "Не позднее даты" },
  { kind: "pile_rhythm", label: "По N свай" },
  { kind: "after_payment", label: "Через N дней после полной оплаты" },
];

const DELIVERY_OPTIONS: { kind: SpecificationDeliveryKind; label: string }[] = [
  { kind: "pickup", label: "Самовывоз" },
  { kind: "site", label: "Доставка на объект" },
];

const SHARE_PERCENTS = [10, 20, 30, 40] as const;

const PAYMENT_DAY_DEFAULTS: Partial<Record<SpecificationPaymentKind, string>> = {
  split_50_50: "3",
  split_share: "5",
  deferral: "45",
};

type Draft = {
  payment: SpecificationPaymentKind | "";
  payment_date: string;
  payment_days: string;
  second_share_percent: number | null;
  second_payment_date: string;
  custom_text: string;
  term: SpecificationTermKind | "";
  term_date: string;
  pile_count: string;
  pile_unit: "week" | "day" | "";
  term_days: string;
  delivery: SpecificationDeliveryKind | "";
  delivery_address: string;
};

type Props = {
  view: SpecificationView;
  busy: boolean;
  error: string | null;
  onSave: (choice: SpecificationChoicePayload) => Promise<void>;
  onDownload: (choice: SpecificationChoicePayload) => Promise<void>;
};

const isPayment = (value: string | null | undefined): value is SpecificationPaymentKind =>
  value === "prepay_100" ||
  value === "split_50_50" ||
  value === "split_share" ||
  value === "deferral" ||
  value === "custom";

const isTerm = (value: string | null | undefined): value is SpecificationTermKind =>
  value === "by_date" || value === "pile_rhythm" || value === "after_payment";

const isDelivery = (value: string | null | undefined): value is SpecificationDeliveryKind =>
  value === "pickup" || value === "site";

const draftFromChoice = (choice: SpecificationChoice | null): Draft => {
  const payment = choice?.payment;
  const term = choice?.term;
  const delivery = choice?.delivery;
  const pileUnit = choice?.pile_unit;
  return {
    payment: isPayment(payment) ? payment : "",
    payment_date: choice?.payment_date ?? "",
    payment_days: choice?.payment_days != null ? String(choice.payment_days) : "",
    second_share_percent: choice?.second_share_percent ?? null,
    second_payment_date: choice?.second_payment_date ?? "",
    custom_text: choice?.custom_text ?? "",
    term: isTerm(term) ? term : "",
    term_date: choice?.term_date ?? "",
    pile_count: choice?.pile_count != null ? String(choice.pile_count) : "",
    pile_unit: pileUnit === "week" || pileUnit === "day" ? pileUnit : "",
    term_days: choice?.term_days != null ? String(choice.term_days) : "",
    delivery: isDelivery(delivery) ? delivery : "",
    delivery_address: choice?.delivery_address ?? "",
  };
};

const optionalInt = (value: string): number | null => {
  if (!value.trim()) {
    return null;
  }
  const parsed = Number(value);
  return Number.isInteger(parsed) ? parsed : null;
};

const toPayload = (draft: Draft): SpecificationChoicePayload | string => {
  if (!draft.payment) {
    return "Выберите схему оплаты";
  }
  if (!draft.term) {
    return "Выберите срок поставки";
  }
  if (!draft.delivery) {
    return "Выберите вид поставки";
  }
  const usesPaymentDays =
    draft.payment === "split_50_50" || draft.payment === "split_share" || draft.payment === "deferral";
  return {
    payment: draft.payment,
    term: draft.term,
    delivery: draft.delivery,
    payment_date: draft.payment_date || null,
    payment_days: usesPaymentDays ? optionalInt(draft.payment_days) : null,
    second_share_percent: draft.payment === "split_share" ? draft.second_share_percent : null,
    second_payment_date: draft.payment === "split_share" ? draft.second_payment_date || null : null,
    custom_text: draft.payment === "custom" ? draft.custom_text.trim() || null : null,
    term_date: draft.term === "after_payment" ? null : draft.term_date || null,
    pile_count: draft.term === "pile_rhythm" ? optionalInt(draft.pile_count) : null,
    pile_unit: draft.term === "pile_rhythm" ? draft.pile_unit || null : null,
    term_days: draft.term === "after_payment" ? optionalInt(draft.term_days) : null,
    delivery_address: draft.delivery === "site" ? draft.delivery_address.trim() || null : null,
  };
};

function ChoiceButton({
  pressed,
  disabled,
  onClick,
  children,
}: {
  pressed: boolean;
  disabled: boolean;
  onClick: () => void;
  children: string;
}) {
  return (
    <Button
      type="button"
      variant={pressed ? "primary" : "secondary"}
      disabled={disabled}
      aria-pressed={pressed}
      onClick={onClick}
    >
      {children}
    </Button>
  );
}

function DateField({
  label,
  value,
  disabled,
  onChange,
}: {
  label: string;
  value: string;
  disabled: boolean;
  onChange: (value: string) => void;
}) {
  return (
    <label style={{ display: "grid", gap: "0.25rem" }}>
      {label}
      <input
        aria-label={label}
        type="date"
        value={value}
        disabled={disabled}
        onChange={(event) => onChange(event.target.value)}
      />
    </label>
  );
}

function DaysField({
  label,
  value,
  disabled,
  onChange,
}: {
  label: string;
  value: string;
  disabled: boolean;
  onChange: (value: string) => void;
}) {
  return (
    <label style={{ display: "grid", gap: "0.25rem" }}>
      {label}
      <input
        aria-label={label}
        type="number"
        min={1}
        max={365}
        value={value}
        disabled={disabled}
        onChange={(event) => onChange(event.target.value)}
      />
    </label>
  );
}

export function SpecificationPanel({ view, busy, error, onSave, onDownload }: Props) {
  const [draft, setDraft] = useState(() => draftFromChoice(view.choice));
  const [localError, setLocalError] = useState<string | null>(null);
  const shownError = localError ?? error;
  const termOptions = TERM_OPTIONS.filter((option) => option.kind !== "pile_rhythm" || view.has_piles);

  const selectPayment = (payment: SpecificationPaymentKind) => {
    setDraft((current) => {
      const previousDefault = current.payment ? PAYMENT_DAY_DEFAULTS[current.payment] : undefined;
      const nextDefault = PAYMENT_DAY_DEFAULTS[payment];
      const keepTyped = Boolean(current.payment_days) && current.payment_days !== previousDefault;
      return {
        ...current,
        payment,
        payment_days: keepTyped ? current.payment_days : nextDefault ?? "",
      };
    });
  };

  const selectTerm = (term: SpecificationTermKind) => {
    setDraft((current) => ({
      ...current,
      term,
      term_days: term === "after_payment" ? current.term_days || "10" : current.term_days,
    }));
  };

  const submit = async (action: (choice: SpecificationChoicePayload) => Promise<void>) => {
    const payload = toPayload(draft);
    if (typeof payload === "string") {
      setLocalError(payload);
      return;
    }
    setLocalError(null);
    await action(payload);
  };

  return (
    <section aria-label="Панель спецификации" style={{ display: "grid", gap: "0.75rem", marginTop: "0.75rem" }}>
      {view.stale_custom && <Alert tone="warning">{STALE_SPECIFICATION}</Alert>}
      {shownError && <Alert tone="error">{shownError}</Alert>}
      <fieldset style={{ display: "grid", gap: "0.5rem" }}>
        <legend>Оплата</legend>
        <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
          {PAYMENT_OPTIONS.map((option) => (
            <ChoiceButton
              key={option.kind}
              pressed={draft.payment === option.kind}
              disabled={busy}
              onClick={() => selectPayment(option.kind)}
            >
              {option.label}
            </ChoiceButton>
          ))}
        </div>
        {(draft.payment === "split_50_50" || draft.payment === "split_share") && (
          <DateField
            label="Дата первого аванса"
            value={draft.payment_date}
            disabled={busy}
            onChange={(payment_date) => setDraft((current) => ({ ...current, payment_date }))}
          />
        )}
        {draft.payment === "split_share" && (
          <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
            {SHARE_PERCENTS.map((percent) => (
              <ChoiceButton
                key={percent}
                pressed={draft.second_share_percent === percent}
                disabled={busy}
                onClick={() => setDraft((current) => ({ ...current, second_share_percent: percent }))}
              >
                {`${percent}%`}
              </ChoiceButton>
            ))}
          </div>
        )}
        {draft.payment === "split_share" && (
          <DateField
            label="Дата второго аванса"
            value={draft.second_payment_date}
            disabled={busy}
            onChange={(second_payment_date) => setDraft((current) => ({ ...current, second_payment_date }))}
          />
        )}
        {draft.payment === "split_50_50" && (
          <DaysField
            label="Дней до отгрузки"
            value={draft.payment_days}
            disabled={busy}
            onChange={(payment_days) => setDraft((current) => ({ ...current, payment_days }))}
          />
        )}
        {draft.payment === "split_share" && (
          <DaysField
            label="Дней после поставки"
            value={draft.payment_days}
            disabled={busy}
            onChange={(payment_days) => setDraft((current) => ({ ...current, payment_days }))}
          />
        )}
        {draft.payment === "deferral" && (
          <DaysField
            label="Дней отсрочки"
            value={draft.payment_days}
            disabled={busy}
            onChange={(payment_days) => setDraft((current) => ({ ...current, payment_days }))}
          />
        )}
        {draft.payment === "custom" && (
          <label style={{ display: "grid", gap: "0.25rem" }}>
            Свой текст
            <textarea
              aria-label="Свой текст"
              rows={4}
              maxLength={2000}
              value={draft.custom_text}
              disabled={busy}
              onChange={(event) => setDraft((current) => ({ ...current, custom_text: event.target.value }))}
            />
          </label>
        )}
        {view.payment_paragraph && <p>{view.payment_paragraph}</p>}
      </fieldset>
      <fieldset style={{ display: "grid", gap: "0.5rem" }}>
        <legend>Срок</legend>
        <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
          {termOptions.map((option) => (
            <ChoiceButton
              key={option.kind}
              pressed={draft.term === option.kind}
              disabled={busy}
              onClick={() => selectTerm(option.kind)}
            >
              {option.label}
            </ChoiceButton>
          ))}
        </div>
        {draft.term === "by_date" && (
          <DateField
            label="Дата поставки"
            value={draft.term_date}
            disabled={busy}
            onChange={(term_date) => setDraft((current) => ({ ...current, term_date }))}
          />
        )}
        {draft.term === "pile_rhythm" && view.has_piles && (
          <>
            <DateField
              label="Дата начала поставок"
              value={draft.term_date}
              disabled={busy}
              onChange={(term_date) => setDraft((current) => ({ ...current, term_date }))}
            />
            <label style={{ display: "grid", gap: "0.25rem" }}>
              Число свай
              <input
                aria-label="Число свай"
                type="number"
                min={1}
                value={draft.pile_count}
                disabled={busy}
                onChange={(event) => setDraft((current) => ({ ...current, pile_count: event.target.value }))}
              />
            </label>
            <div style={{ display: "flex", gap: "0.5rem" }}>
              <ChoiceButton
                pressed={draft.pile_unit === "week"}
                disabled={busy}
                onClick={() => setDraft((current) => ({ ...current, pile_unit: "week" }))}
              >
                в неделю
              </ChoiceButton>
              <ChoiceButton
                pressed={draft.pile_unit === "day"}
                disabled={busy}
                onClick={() => setDraft((current) => ({ ...current, pile_unit: "day" }))}
              >
                в день
              </ChoiceButton>
            </div>
          </>
        )}
        {draft.term === "after_payment" && (
          <DaysField
            label="Дней после оплаты"
            value={draft.term_days}
            disabled={busy}
            onChange={(term_days) => setDraft((current) => ({ ...current, term_days }))}
          />
        )}
        {view.has_piles && view.concrete_grade && <p>Марка бетона: {view.concrete_grade}</p>}
        {view.term_paragraph && <p>{view.term_paragraph}</p>}
      </fieldset>
      <fieldset style={{ display: "grid", gap: "0.5rem" }}>
        <legend>Вид поставки</legend>
        <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
          {DELIVERY_OPTIONS.map((option) => (
            <ChoiceButton
              key={option.kind}
              pressed={draft.delivery === option.kind}
              disabled={busy}
              onClick={() => setDraft((current) => ({ ...current, delivery: option.kind }))}
            >
              {option.label}
            </ChoiceButton>
          ))}
        </div>
        {draft.delivery === "site" && (
          <label style={{ display: "grid", gap: "0.25rem" }}>
            Адрес объекта
            <input
              aria-label="Адрес объекта"
              maxLength={300}
              value={draft.delivery_address}
              disabled={busy}
              onChange={(event) => setDraft((current) => ({ ...current, delivery_address: event.target.value }))}
            />
          </label>
        )}
        {view.delivery_paragraph && <p>{view.delivery_paragraph}</p>}
      </fieldset>
      <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
        <Button type="button" variant="primary" disabled={busy} onClick={() => void submit(onSave)}>
          Сохранить
        </Button>
        <Button type="button" variant="secondary" disabled={busy} onClick={() => void submit(onDownload)}>
          Скачать Excel
        </Button>
      </div>
    </section>
  );
}

export function specificationPanelKey(view: SpecificationView): string {
  return [
    view.saved,
    view.stale_custom,
    view.payment_paragraph,
    view.term_paragraph,
    view.delivery_paragraph,
    view.spec_date ?? "",
  ].join("|");
}
