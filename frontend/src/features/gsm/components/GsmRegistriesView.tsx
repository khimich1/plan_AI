import type { CSSProperties } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";
import { formatGsmError } from "@/features/gsm/lib/gsmErrors";
import { CardsRegistryView } from "@/features/gsm/components/CardsRegistryView";
import { DriversRegistryView } from "@/features/gsm/components/DriversRegistryView";
import { VehiclesCard } from "@/features/gsm/components/VehiclesCard";
import {
  useGsmSettingsQuery,
  useGsmStationsQuery,
  useUpdateGsmSeasonMutation,
} from "@/features/gsm/hooks/useGsmQueries";
import type { GsmSettings } from "@/features/gsm/types/gsm";

const sectionStyle: CSSProperties = {
  display: "grid",
  gap: "0.75rem",
  padding: "1rem",
  borderRadius: 12,
  border: "1px solid #e4e7ec",
  background: "#ffffff",
};

const thStyle: CSSProperties = { padding: "0.5rem", textAlign: "left", borderBottom: "1px solid #eaecf0" };
const tdStyle: CSSProperties = { padding: "0.5rem", borderBottom: "1px solid #f2f4f7" };

const todayIsoDate = (): string => {
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${now.getFullYear()}-${month}-${day}`;
};

const seasonLabel = (settings: GsmSettings): string => {
  const mode = settings.season_mode === "winter" ? "зимний" : "летний";
  return settings.season_switched_at ? `${mode} (с ${settings.season_switched_at})` : mode;
};

export const GsmRegistriesView = () => {
  const stationsQuery = useGsmStationsQuery();
  const settingsQuery = useGsmSettingsQuery();
  const seasonMutation = useUpdateGsmSeasonMutation();

  const isBootLoading = stationsQuery.isLoading || settingsQuery.isLoading;
  const bootError = stationsQuery.error ?? settingsQuery.error;

  if (isBootLoading) {
    return <Spinner />;
  }

  if (bootError) {
    return <Alert tone="error">{formatGsmError(bootError)}</Alert>;
  }

  const stations = stationsQuery.data ?? [];
  const settings = settingsQuery.data;
  const targetMode = settings?.season_mode === "winter" ? "summer" : "winter";

  return (
    <section style={{ display: "grid", gap: "1.25rem" }} aria-label="Справочники ГСМ">
      {settings && (
        <div style={sectionStyle}>
          <h2 style={{ margin: 0, fontSize: "1.1rem" }}>Настройки</h2>
          <p style={{ margin: 0, color: "#475467" }}>
            Режим: {seasonLabel(settings)}, порог крюка {settings.hook_threshold_km} км
          </p>
          <div>
            <Button
              type="button"
              variant="secondary"
              disabled={seasonMutation.isPending}
              onClick={() => seasonMutation.mutate({ mode: targetMode, date: todayIsoDate() })}
            >
              {targetMode === "winter" ? "Перевести на зимний режим" : "Перевести на летний режим"}
            </Button>
          </div>
          {seasonMutation.isError && (
            <Alert tone="error">{formatGsmError(seasonMutation.error)}</Alert>
          )}
        </div>
      )}

      <VehiclesCard />
      <DriversRegistryView />
      <CardsRegistryView />

      <div style={sectionStyle}>
        <h2 style={{ margin: 0, fontSize: "1.1rem" }}>АЗС ({stations.length})</h2>
        {stations.length === 0 ? (
          <Alert tone="info">Станции не найдены.</Alert>
        ) : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.92rem" }}>
              <thead>
                <tr>
                  <th style={thStyle}>Адрес</th>
                  <th style={thStyle}>Бренд</th>
                </tr>
              </thead>
              <tbody>
                {stations.map((s) => (
                  <tr key={s.id}>
                    <td style={tdStyle}>{s.address}</td>
                    <td style={tdStyle}>{s.brand ?? "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </section>
  );
};
