import { useState, type CSSProperties } from "react";
import type { PlanIntegrityReport } from "@/features/production/types/production";

type Props = {
  report?: PlanIntegrityReport | null;
};

const chipBase: CSSProperties = {
  fontSize: "0.75rem",
  fontWeight: 600,
  padding: "0.15rem 0.5rem",
  borderRadius: 999,
  lineHeight: 1.4,
  border: "none",
  fontFamily: "inherit",
};

const listStyle: CSSProperties = {
  margin: "0.35rem 0 0",
  padding: "0.4rem 0.75rem",
  listStyle: "none",
  background: "#ffffff",
  border: "1px solid #e4e7ec",
  borderRadius: 8,
  color: "#344054",
  fontSize: "0.8rem",
  fontWeight: 500,
  display: "grid",
  gap: "0.2rem",
};

function discrepancyTotal(report: PlanIntegrityReport): number {
  return report.orphans + report.surplus + report.no_grade;
}

export const PlanHealthBadge = ({ report }: Props) => {
  const [open, setOpen] = useState(false);
  if (!report) {
    return null;
  }

  const total = discrepancyTotal(report);
  if (total <= 0) {
    return (
      <span
        style={{
          ...chipBase,
          background: "#ecfdf3",
          color: "#067647",
        }}
      >
        чист
      </span>
    );
  }

  return (
    <span style={{ display: "inline-flex", flexDirection: "column", alignItems: "flex-start" }}>
      <button
        type="button"
        aria-expanded={open}
        aria-label="Расхождения плана"
        onClick={() => setOpen((value) => !value)}
        style={{
          ...chipBase,
          background: "#fffaeb",
          color: "#b54708",
          cursor: "pointer",
        }}
      >
        {`${total} расхождений`}
      </button>
      {open && (
        <ul style={listStyle}>
          {report.items.map((item, index) => (
            <li key={`${item.kind}-${index}`}>{item.message}</li>
          ))}
        </ul>
      )}
    </span>
  );
};
