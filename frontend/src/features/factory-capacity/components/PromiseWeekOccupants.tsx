import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import {
  formatQuoteDayMonth,
  usePromiseWeekOccupantsQuery,
  type PromiseWeekOccupant,
} from "@/features/factory-capacity/api/promiseQuote";
import { getErrorMessage } from "@/shared/lib/apiError";
import type { CSSProperties } from "react";

type Props = {
  kpId: number;
  weekStart: string | null;
  weekFree?: number;
  weekCapacity?: number;
  selectedDay?: string | null;
  occupancy?: Record<string, number>;
  knob?: number;
  holidays?: string[];
  extraWorkdays?: string[];
};

const KIND_LABEL: Record<PromiseWeekOccupant["kind"], string> = {
  hold: "холд",
  promise: "обещано",
};

const wrapStyle: CSSProperties = {
  display: "grid",
  gap: "0.6rem",
  marginTop: "0.75rem",
};

const listStyle: CSSProperties = {
  display: "grid",
  gap: "0.4rem",
  margin: 0,
  padding: 0,
  listStyle: "none",
};

const rowStyle = (current: boolean): CSSProperties => ({
  display: "grid",
  gap: "0.15rem",
  padding: "0.55rem 0.7rem",
  borderRadius: 10,
  border: current ? "1px solid #2e90fa" : "1px solid #e4e7ec",
  background: current ? "#f5f9ff" : "#ffffff",
  fontSize: "0.85rem",
});

function isWeekend(iso: string): boolean {
  const day = new Date(`${iso}T12:00:00`).getDay();
  return day === 0 || day === 6;
}

function isNonWorking(iso: string, holidays: string[], extraWorkdays: string[]): boolean {
  if (extraWorkdays.includes(iso)) return false;
  if (holidays.includes(iso)) return true;
  return isWeekend(iso);
}

function selectedDayLine(
  iso: string,
  occupancy: Record<string, number>,
  knob: number,
  holidays: string[],
  extraWorkdays: string[],
): string {
  const label = formatQuoteDayMonth(iso);
  if (isNonWorking(iso, holidays, extraWorkdays)) {
    return `${label}: нерабочий`;
  }
  const occupied = occupancy[iso] ?? 0;
  if (occupied <= 0) {
    return `${label}: свободно ${knob}`;
  }
  if (occupied > knob) {
    return `${label}: ${occupied}/${knob} · перебор`;
  }
  return `${label}: ${occupied}/${knob}, остаток ${Math.max(0, knob - occupied)}`;
}

export const PromiseWeekOccupants = ({
  kpId,
  weekStart,
  weekFree,
  weekCapacity,
  selectedDay,
  occupancy = {},
  knob,
  holidays = [],
  extraWorkdays = [],
}: Props) => {
  const query = usePromiseWeekOccupantsQuery(kpId, weekStart);

  return (
    <section data-testid="promise-week-occupants" style={wrapStyle}>
      {query.isPending ? (
        <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
          <Spinner /> Загружаю жильцов недели…
        </div>
      ) : null}
      {query.isError ? <Alert tone="error">{getErrorMessage(query.error)}</Alert> : null}
      {query.data ? (
        <OccupantsBody
          data={query.data}
          weekFree={weekFree}
          weekCapacity={weekCapacity}
          selectedDay={selectedDay}
          occupancy={occupancy}
          knob={knob}
          holidays={holidays}
          extraWorkdays={extraWorkdays}
        />
      ) : null}
    </section>
  );
};

function OccupantsBody({
  data,
  weekFree,
  weekCapacity,
  selectedDay,
  occupancy,
  knob,
  holidays,
  extraWorkdays,
}: {
  data: { planned: number; occupants: PromiseWeekOccupant[] };
  weekFree?: number;
  weekCapacity?: number;
  selectedDay?: string | null;
  occupancy: Record<string, number>;
  knob?: number;
  holidays: string[];
  extraWorkdays: string[];
}) {
  return (
    <>
      <div data-testid="promise-week-planned" style={{ fontSize: "0.9rem", color: "#101828" }}>
        Уже в плане: {data.planned} дорожек
      </div>
      {weekFree !== undefined && weekCapacity !== undefined ? (
        <div data-testid="promise-week-free" style={{ fontSize: "0.9rem", color: "#101828" }}>
          Свободно: {weekFree} из {weekCapacity}
        </div>
      ) : null}
      {selectedDay && knob !== undefined ? (
        <div data-testid="promise-week-day-line" style={{ fontSize: "0.9rem", color: "#475467" }}>
          {selectedDayLine(selectedDay, occupancy, knob, holidays, extraWorkdays)}
        </div>
      ) : null}
      {data.occupants.length === 0 ? (
        <p data-testid="promise-week-occupants-empty" style={{ margin: 0, color: "#667085" }}>
          На этой неделе нет холдов и обещаний
        </p>
      ) : (
        <ul style={listStyle}>
          {data.occupants.map((row) => (
            <li
              key={`${row.kind}-${row.kp_id}`}
              data-testid="promise-week-occupant"
              data-current={row.is_current ? "true" : "false"}
              style={rowStyle(row.is_current)}
            >
              <strong>
                КП №{row.kp_id}
                {row.is_current ? " · это КП" : ""}
              </strong>
              <span style={{ color: "#475467" }}>{row.customer_name.trim() || "—"}</span>
              <span style={{ color: "#667085" }}>
                {KIND_LABEL[row.kind]} · {row.tracks} дор. · к {formatQuoteDayMonth(row.promised_date)}
              </span>
            </li>
          ))}
        </ul>
      )}
      <p data-testid="promise-week-hold-caption" style={{ margin: 0, fontSize: "0.8rem", color: "#667085" }}>
        Холды не занимают свободно — до перевода место могут взять другие.
      </p>
    </>
  );
}
