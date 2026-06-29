import { Button } from "@/shared/ui/Button";
import { FieldWrapper, Input } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import { MonthCalendarGrid } from "@/features/production/components/MonthCalendarGrid";
import type { DayInfo } from "@/features/production/types/production";

type Props = {
  startDate: string;
  planName: string;
  calendarMonth: Date;
  daysInfo: Record<string, DayInfo>;
  holidays: Set<string>;
  extraWorkdays: Set<string>;
  occupiedOnStart: number;
  maxPerDay: number;
  freeOnStart: number;
  calendarLoading: boolean;
  canProceed: boolean;
  onStartDateChange: (date: string) => void;
  onPlanNameChange: (name: string) => void;
  onCalendarMonthChange: (month: Date) => void;
  onNext: () => void;
};

export const Step1PlanStartDate = ({
  startDate,
  planName,
  calendarMonth,
  daysInfo,
  holidays,
  extraWorkdays,
  occupiedOnStart,
  maxPerDay,
  freeOnStart,
  calendarLoading,
  canProceed,
  onStartDateChange,
  onPlanNameChange,
  onCalendarMonthChange,
  onNext,
}: Props) => (
  <div style={{ display: "grid", gap: "1rem" }}>
    <FieldWrapper
      label="Дата начала планирования"
      hint="План построится начиная с этого дня (рабочие дни учитываются автоматически)."
    >
      <Input
        type="date"
        value={startDate}
        onChange={(e) => onStartDateChange(e.target.value)}
      />
    </FieldWrapper>

    <FieldWrapper label="Название плана (необязательно)">
      <Input
        type="text"
        placeholder="Например: «План на 20.04»"
        value={planName}
        onChange={(e) => onPlanNameChange(e.target.value)}
      />
    </FieldWrapper>

    <div>
      <div style={{ fontWeight: 600, marginBottom: "0.5rem" }}>
        Загрузка дат (из других планов):
      </div>
      {calendarLoading ? (
        <Spinner />
      ) : (
        <MonthCalendarGrid
          daysInfo={daysInfo}
          holidays={holidays}
          extraWorkdays={extraWorkdays}
          month={calendarMonth}
          onMonthChange={onCalendarMonthChange}
          selectedDate={startDate}
          onSelectDate={onStartDateChange}
        />
      )}
      <div style={{ color: "#475467", marginTop: "0.35rem" }}>
        На выбранный день уже занято <strong>{occupiedOnStart}</strong> из{" "}
        {maxPerDay} дорожек (свободно {freeOnStart}).
      </div>
    </div>

    <div style={{ display: "flex", justifyContent: "flex-end" }}>
      <Button onClick={onNext} disabled={!canProceed}>
        Далее →
      </Button>
    </div>
  </div>
);
