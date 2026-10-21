import type { CSSProperties } from "react";
import type { TimesheetGridCell } from "@/features/gsm/types/gsm";

type Props = {
  period: string;
  cells: TimesheetGridCell[];
};

const thStyle: CSSProperties = {
  padding: "0.35rem 0.4rem",
  textAlign: "center",
  borderBottom: "1px solid #eaecf0",
  fontSize: "0.75rem",
  color: "#475467",
  position: "sticky",
  top: 0,
  background: "#fff",
};

const nameStyle: CSSProperties = {
  padding: "0.35rem 0.5rem",
  borderBottom: "1px solid #f2f4f7",
  whiteSpace: "nowrap",
  position: "sticky",
  left: 0,
  background: "#fff",
};

const daysInMonth = (period: string): number => {
  const [yearText, monthText] = period.split("-");
  const year = Number(yearText);
  const month = Number(monthText);
  if (!year || !month) {
    return 31;
  }
  return new Date(year, month, 0).getDate();
};

export const TimesheetGridView = ({ period, cells }: Props) => {
  const last = daysInMonth(period);
  const byDriver = new Map<number, { name: string; codes: Map<number, string> }>();
  for (const cell of cells) {
    const day = Number(cell.date.slice(8, 10));
    const row = byDriver.get(cell.driver_id) ?? { name: cell.full_name, codes: new Map() };
    row.codes.set(day, cell.code);
    byDriver.set(cell.driver_id, row);
  }
  const drivers = [...byDriver.entries()].sort((a, b) => a[1].name.localeCompare(b[1].name, "ru"));

  if (drivers.length === 0) {
    return <p style={{ margin: 0, color: "#475467" }}>За {period} явок нет.</p>;
  }

  return (
    <div style={{ overflow: "auto", maxHeight: 420 }} aria-label="Сетка табеля">
      <table style={{ borderCollapse: "collapse", fontSize: "0.8rem" }}>
        <thead>
          <tr>
            <th style={{ ...thStyle, textAlign: "left", left: 0, zIndex: 1 }}>Водитель</th>
            {Array.from({ length: last }, (_, index) => (
              <th key={index + 1} style={thStyle}>
                {index + 1}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {drivers.map(([driverId, row]) => (
            <tr key={driverId}>
              <th scope="row" style={nameStyle}>
                {row.name}
              </th>
              {Array.from({ length: last }, (_, index) => {
                const day = index + 1;
                const code = row.codes.get(day) ?? "";
                const absent = code === "ОТ" || (code !== "" && code !== "Я" && code !== "В");
                const isoDay = String(day).padStart(2, "0");
                return (
                  <td
                    key={day}
                    data-testid={`timesheet-cell-${driverId}-${period}-${isoDay}`}
                    data-absent={absent ? "true" : "false"}
                    style={{
                      padding: "0.3rem",
                      textAlign: "center",
                      borderBottom: "1px solid #f2f4f7",
                      background: absent ? "#fef3f2" : undefined,
                      color: absent ? "#b42318" : "#101828",
                      fontWeight: absent ? 700 : 500,
                    }}
                  >
                    {code || "—"}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};
