import { CarriersView } from "@/features/logistics/components/CarriersView";
import { LogisticsTabs } from "@/features/logistics/components/LogisticsTabs";

export const LogisticsCarriersPage = () => (
  <main style={{ maxWidth: 1280, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
    <div style={{ display: "grid", gap: "1rem" }}>
      <header>
        <h1 style={{ margin: 0, fontSize: "1.75rem" }}>Логистика</h1>
        <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
          Справочник перевозчиков: поиск и слияние дублей.
        </p>
      </header>
      <div>
        <LogisticsTabs />
      </div>
      <CarriersView />
    </div>
  </main>
);
