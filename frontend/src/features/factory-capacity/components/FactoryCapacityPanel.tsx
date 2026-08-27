import type { CSSProperties } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import { FactoryMiniCalendar } from "@/features/factory-capacity/components/FactoryMiniCalendar";
import type { CapacitySnapshot, CapacityStatus } from "@/features/factory-capacity/types/capacity";

type Props = {
  snapshot: CapacitySnapshot | null | undefined;
  isLoading?: boolean;
  errorMessage?: string | null;
};

const statusTone: Record<CapacityStatus, CSSProperties> = {
  green: { color: "#067647" },
  yellow: { color: "#b54708" },
  red: { color: "#b42318" },
};

export const FactoryCapacityPanel = ({ snapshot, isLoading, errorMessage }: Props) => {
  if (isLoading) {
    return (
      <div
        data-testid="factory-capacity-panel"
        style={{ display: "flex", gap: "0.5rem", alignItems: "center", fontSize: "0.9rem" }}
      >
        <Spinner /> Считаю загрузку завода…
      </div>
    );
  }

  if (errorMessage) {
    return (
      <div data-testid="factory-capacity-panel">
        <Alert tone="warning">{errorMessage}</Alert>
      </div>
    );
  }

  if (!snapshot) {
    return null;
  }

  const tone = statusTone[snapshot.status];

  return (
    <aside
      data-testid="factory-capacity-panel"
      style={{
        display: "grid",
        gap: "0.75rem",
        padding: "0.85rem 1rem",
        border: "1px solid #e4e7ec",
        borderRadius: 12,
        background: "#fafafa",
        minWidth: 260,
      }}
    >
      <div style={{ display: "grid", gap: "0.25rem" }}>
        <strong style={{ fontSize: "0.85rem" }}>Ёмкость завода</strong>
        <div style={{ fontSize: "0.85rem", display: "flex", flexWrap: "wrap", gap: "0.75rem" }}>
          <span>
            нужно <strong>{snapshot.tracks_needed}</strong>
          </span>
          <span>
            свободно <strong>{snapshot.tracks_free_in_window}</strong>
          </span>
          <span style={tone}>
            Δ <strong>{snapshot.delta}</strong>
          </span>
        </div>
      </div>

      <FactoryMiniCalendar snapshot={snapshot} />

      {snapshot.status === "red" && snapshot.hint ? (
        <Alert tone="error">{snapshot.hint}. Увеличьте срок в поле выше и повторите.</Alert>
      ) : null}
      {snapshot.status === "yellow" ? (
        <Alert tone="warning">Срок на грани — сохранить можно, но запас небольшой.</Alert>
      ) : null}
    </aside>
  );
};

/** True when save/submit must be blocked by capacity gate. */
export const isCapacityRed = (snapshot: CapacitySnapshot | null | undefined): boolean =>
  snapshot?.status === "red";
