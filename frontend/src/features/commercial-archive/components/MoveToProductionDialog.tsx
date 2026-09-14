import { useEffect, useMemo, useRef, useState } from "react";
import { useQueryClient } from "@tanstack/react-query";
import { Modal } from "@/shared/ui/Modal";
import { Drawer } from "@/shared/ui/Drawer";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import { useMoveToProductionMutation } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { PromiseKnobSettings } from "@/features/factory-capacity/components/PromiseKnobSettings";
import { PromisePeriodCalendar } from "@/features/factory-capacity/components/PromisePeriodCalendar";
import { PromiseQuoteBlock } from "@/features/factory-capacity/components/PromiseQuoteBlock";
import { PromiseWeekOccupants } from "@/features/factory-capacity/components/PromiseWeekOccupants";
import { PromiseWindowBand } from "@/features/factory-capacity/components/PromiseWindowBand";
import {
  addDaysIso,
  holdCreatedByTitle,
  isoToDdMmYyyy,
  promiseHoldKeys,
  promiseQuoteKeys,
  useCreatePromiseHoldMutation,
  usePromiseHoldQuery,
  usePromiseQuoteQuery,
} from "@/features/factory-capacity/api/promiseQuote";
import { ddMmYyyyToIso } from "@/features/factory-capacity/lib/dates";
import { isoWeekStart } from "@/features/factory-capacity/lib/isoWeek";
import { getErrorMessage } from "@/shared/lib/apiError";
import {
  EXECUTION_TERMS_FIELD_HINT,
  EXECUTION_TERMS_PARSE_ERROR,
  EXECUTION_TERMS_PLACEHOLDER,
  tryNormalizeExecutionTerms,
} from "@/shared/lib/executionTerms";

type Props = {
  open: boolean;
  onClose: () => void;
  kpId: number;
  /** Уже сохранённый в карточке КП срок (обычно ДД.ММ.ГГГГ) — показываем как стартовое значение. */
  initialExecutionTerms?: string | null;
};

const DRAWER_WIDTH = 420;

function firstOfMonth(iso: string): string {
  return `${iso.slice(0, 7)}-01`;
}

function maxMonthIso(left: string, right: string): string {
  return left >= right ? left : right;
}

export const MoveToProductionDialog = ({
  open,
  onClose,
  kpId,
  initialExecutionTerms,
}: Props) => {
  const [value, setValue] = useState("");
  const [capacityOpen, setCapacityOpen] = useState(false);
  const [selectedWeek, setSelectedWeek] = useState<string | null>(null);
  const [selectedDay, setSelectedDay] = useState<string | null>(null);
  const [calendarMonth, setCalendarMonth] = useState(() => firstOfMonth(new Date().toISOString().slice(0, 10)));
  const userEditedRef = useRef(false);
  const queryClient = useQueryClient();
  const mutation = useMoveToProductionMutation();
  const holdMutation = useCreatePromiseHoldMutation();
  const quote = usePromiseQuoteQuery(open ? kpId : null);
  const holdQuery = usePromiseHoldQuery(open ? kpId : null);
  const activeHold = holdQuery.data ?? null;

  useEffect(() => {
    if (!open) {
      userEditedRef.current = false;
      setCapacityOpen(false);
      setSelectedWeek(null);
      setSelectedDay(null);
      return;
    }
    mutation.reset();
    holdMutation.reset();
    userEditedRef.current = false;
    setValue("");
    setCapacityOpen(false);
    setSelectedWeek(null);
    setSelectedDay(null);
  }, [open, kpId]); // eslint-disable-line react-hooks/exhaustive-deps -- только open/kpId: сброс при открытии/смене КП

  useEffect(() => {
    if (!open || userEditedRef.current) {
      return;
    }
    const fromHold = isoToDdMmYyyy(activeHold?.promised_date);
    if (fromHold) {
      setValue(fromHold);
      return;
    }

    const fromCard = (initialExecutionTerms ?? "").trim();
    if (fromCard) {
      setValue(fromCard);
      return;
    }

    if (quote.isPending) {
      setValue("");
      return;
    }

    const promised = isoToDdMmYyyy(quote.data?.window?.promised_date);
    if (promised) {
      setValue(promised);
      return;
    }

    setValue("");
  }, [
    open,
    kpId,
    initialExecutionTerms,
    quote.isPending,
    quote.data?.window?.promised_date,
    activeHold?.promised_date,
  ]);

  const normalizedTerms = useMemo(() => {
    const trimmed = value.trim();
    if (!trimmed) return null;
    return tryNormalizeExecutionTerms(trimmed);
  }, [value]);

  const targetIso = useMemo(
    () => (normalizedTerms ? ddMmYyyyToIso(normalizedTerms) : null),
    [normalizedTerms],
  );

  const defaultWeek = useMemo(() => {
    const promised = quote.data?.window?.promised_date;
    if (promised) return isoWeekStart(promised);
    return quote.data?.weeks[0]?.week_start ?? null;
  }, [quote.data?.window?.promised_date, quote.data?.weeks]);

  const activeWeekStart = selectedWeek ?? defaultWeek;
  const activeWeek = quote.data?.weeks.find((week) => week.week_start === activeWeekStart);
  const pourToSunday = quote.data?.solo_date
    ? addDaysIso(isoWeekStart(quote.data.solo_date), 6)
    : null;

  useEffect(() => {
    if (!open || !defaultWeek) {
      return;
    }
    setSelectedWeek((prev) => prev ?? defaultWeek);
  }, [open, defaultWeek]);

  const fieldMonth = targetIso ? firstOfMonth(targetIso) : null;

  const desiredMonth = useMemo(() => {
    const promised = quote.data?.window?.promised_date;
    const base = firstOfMonth(
      promised ?? defaultWeek ?? new Date().toISOString().slice(0, 10),
    );
    if (fieldMonth && fieldMonth > base) {
      return fieldMonth;
    }
    return base;
  }, [quote.data?.window?.promised_date, defaultWeek, fieldMonth]);

  useEffect(() => {
    if (!open) {
      return;
    }
    setCalendarMonth(desiredMonth);
  }, [open, desiredMonth]);

  const minMonth = useMemo(() => {
    const firstWeek = quote.data?.weeks[0]?.week_start;
    return firstWeek ? firstOfMonth(firstWeek) : firstOfMonth(new Date().toISOString().slice(0, 10));
  }, [quote.data?.weeks]);

  const maxMonth = useMemo(() => {
    const weeks = quote.data?.weeks ?? [];
    const lastWeek = weeks[weeks.length - 1]?.week_start;
    let end = minMonth;
    if (lastWeek) {
      const sunday = new Date(`${lastWeek}T12:00:00`);
      sunday.setDate(sunday.getDate() + 6);
      end = firstOfMonth(
        `${sunday.getFullYear()}-${String(sunday.getMonth() + 1).padStart(2, "0")}-01`,
      );
    }
    if (fieldMonth) {
      end = maxMonthIso(end, fieldMonth);
    }
    return end;
  }, [quote.data?.weeks, fieldMonth, minMonth]);

  const onSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    const trimmed = value.trim();
    if (!trimmed || tryNormalizeExecutionTerms(trimmed) === null) return;

    try {
      await mutation.mutateAsync({ kpId, executionTerms: trimmed });
      await queryClient.invalidateQueries({ queryKey: promiseQuoteKeys.all });
      await queryClient.invalidateQueries({ queryKey: promiseHoldKeys.all });
      onClose();
    } catch {
      // ошибка показывается через mutation.error
    }
  };

  const onHold = async () => {
    if (!quote.data?.window || activeHold) {
      return;
    }
    try {
      await holdMutation.mutateAsync(kpId);
    } catch {
      // ошибка показывается через holdMutation.error
    }
  };

  const localParseError =
    value.trim().length > 0 && tryNormalizeExecutionTerms(value.trim()) === null
      ? EXECUTION_TERMS_PARSE_ERROR
      : null;

  const handleModalClose = () => {
    if (capacityOpen) {
      setCapacityOpen(false);
      return;
    }
    onClose();
  };

  return (
    <>
      <Modal open={open} onClose={handleModalClose} title={`В производство: КП №${kpId}`}>
        <form onSubmit={onSubmit} style={{ display: "grid", gap: "1rem" }}>
          {quote.isPending && (
            <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
              <Spinner /> Считаю срок по корзинам…
            </div>
          )}
          {quote.isError ? (
            <Alert tone="error">{getErrorMessage(quote.error)}</Alert>
          ) : null}
          {quote.data ? (
            <>
              <PromiseQuoteBlock quote={quote.data} />
              <PromiseWindowBand
                window={quote.data.window}
                firstPourDate={quote.data.first_pour_date}
                pourToSunday={pourToSunday}
              />
            </>
          ) : null}
          {activeHold ? (
            <div
              data-testid="promise-hold-locked"
              role="status"
              title={holdCreatedByTitle(activeHold.created_by)}
              style={{
                border: "1px solid #bfd4ff",
                background: "#eef4ff",
                color: "#1d4ed8",
                borderRadius: 14,
                padding: "0.9rem 1rem",
                fontSize: "0.9rem",
                fontWeight: 600,
              }}
            >
              Срок закреплён до сегодня
              {activeHold.created_by ? ` · ${activeHold.created_by}` : ""}
            </div>
          ) : null}
          <FieldWrapper label="Срок выполнения" hint={EXECUTION_TERMS_FIELD_HINT}>
            <Input
              type="text"
              value={value}
              onChange={(event) => {
                userEditedRef.current = true;
                setValue(event.target.value);
              }}
              placeholder={EXECUTION_TERMS_PLACEHOLDER}
              autoFocus
            />
          </FieldWrapper>
          {localParseError && <Alert tone="error">{localParseError}</Alert>}
          {mutation.isError && <Alert tone="error">{getErrorMessage(mutation.error)}</Alert>}
          {holdMutation.isError && <Alert tone="error">{getErrorMessage(holdMutation.error)}</Alert>}

          <div style={{ display: "flex", justifyContent: "flex-end", gap: "0.5rem", flexWrap: "wrap" }}>
            <Button
              variant="ghost"
              type="button"
              onClick={() => setCapacityOpen(true)}
              disabled={!quote.data}
            >
              Ёмкость
            </Button>
            <Button
              variant="ghost"
              type="button"
              onClick={() => {
                void onHold();
              }}
              disabled={
                holdMutation.isPending ||
                mutation.isPending ||
                Boolean(activeHold) ||
                !quote.data?.window
              }
            >
              {holdMutation.isPending ? "Закрепляю…" : "Закрепить срок"}
            </Button>
            <Button variant="ghost" type="button" onClick={handleModalClose} disabled={mutation.isPending}>
              Отмена
            </Button>
            <Button
              type="submit"
              disabled={
                mutation.isPending ||
                !value.trim() ||
                tryNormalizeExecutionTerms(value.trim()) === null
              }
            >
              {mutation.isPending ? "Перевод..." : "Перевести в производство"}
            </Button>
          </div>
        </form>
      </Modal>

      <Drawer
        open={open && capacityOpen}
        onClose={() => setCapacityOpen(false)}
        title="Ёмкость завода"
        side="left"
        width={DRAWER_WIDTH}
      >
        {quote.data ? (
          <div style={{ display: "grid", gap: "0.75rem" }}>
            <PromisePeriodCalendar
              month={calendarMonth}
              minMonth={minMonth}
              maxMonth={maxMonth}
              onMonthChange={setCalendarMonth}
              selectedWeekStart={activeWeekStart}
              onSelectWeek={setSelectedWeek}
              onSelectDay={setSelectedDay}
              promisedDate={quote.data.window?.promised_date}
              firstPourDate={quote.data.first_pour_date}
              pourFrom={quote.data.first_pour_date}
              pourToSunday={pourToSunday}
              occupancy={quote.data.occupancy ?? {}}
              knob={quote.data.knob}
              holidays={quote.data.holidays ?? []}
              extraWorkdays={quote.data.extra_workdays ?? []}
            />
            <PromiseWeekOccupants
              kpId={kpId}
              weekStart={activeWeekStart}
              weekFree={activeWeek?.free}
              weekCapacity={activeWeek?.capacity}
              selectedDay={selectedDay}
              occupancy={quote.data.occupancy ?? {}}
              knob={quote.data.knob}
              holidays={quote.data.holidays ?? []}
              extraWorkdays={quote.data.extra_workdays ?? []}
            />
            <PromiseKnobSettings currentKnob={quote.data.knob} />
          </div>
        ) : null}
      </Drawer>
    </>
  );
};
