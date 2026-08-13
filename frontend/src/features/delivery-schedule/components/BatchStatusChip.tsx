import type { CSSProperties } from "react";
import type { TrafficLightStatus } from "@/features/delivery-schedule/types/deliverySchedule";

type Props = {
  status: TrafficLightStatus | null | undefined;
  hint?: string | null;
  readyDate?: string | null;
};

const styles: Record<TrafficLightStatus, CSSProperties> = {
  green: { background: "#ecfdf3", color: "#067647", borderColor: "#abefc6" },
  yellow: { background: "#fffaeb", color: "#b54708", borderColor: "#fedf89" },
  red: { background: "#fef3f2", color: "#b42318", borderColor: "#fecdca" },
};

const labels: Record<TrafficLightStatus, string> = {
  green: "В срок",
  yellow: "На грани",
  red: "Риск срыва",
};

export const BatchStatusChip = ({ status, hint, readyDate }: Props) => {
  if (!status) {
    return (
      <span style={{ fontSize: "0.8rem", color: "#667085" }} title="Светофор появится после сохранения">
        нет статуса
      </span>
    );
  }

  const tone = styles[status];
  const titleParts = [labels[status]];
  if (readyDate) {
    titleParts.push(`готовность ~ ${readyDate}`);
  }
  if (hint) {
    titleParts.push(hint);
  }

  return (
    <span
      title={titleParts.join(" · ")}
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: "0.35rem",
        fontSize: "0.75rem",
        fontWeight: 600,
        padding: "0.2rem 0.55rem",
        borderRadius: 999,
        border: `1px solid ${tone.borderColor}`,
        background: tone.background,
        color: tone.color,
        maxWidth: "100%",
      }}
    >
      <span
        aria-hidden
        style={{
          width: 8,
          height: 8,
          borderRadius: "50%",
          background: tone.color,
          flexShrink: 0,
        }}
      />
      <span>{labels[status]}</span>
      {hint ? (
        <span style={{ fontWeight: 500, opacity: 0.9, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          — {hint}
        </span>
      ) : null}
    </span>
  );
};
