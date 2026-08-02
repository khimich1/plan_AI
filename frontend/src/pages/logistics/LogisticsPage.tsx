import { LogisticsTabs } from "@/features/logistics/components/LogisticsTabs";
import { LogisticsRegistryView } from "@/features/logistics/components/LogisticsRegistryView";

export const LogisticsPage = () => (
  <main style={{ maxWidth: 1280, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
    <div style={{ display: "grid", gap: "1rem" }}>
      <header>
        <h1 style={{ margin: 0, fontSize: "1.75rem" }}>Логистика</h1>
        <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
          Реестр рейсов отгрузки: сбор состава из СГП, документы, выезд.
        </p>
      </header>
      <div>
        <LogisticsTabs />
      </div>
      <LogisticsRegistryView />
    </div>
  </main>
);
