import { useEffect } from "react";
import { useSearchParams } from "react-router-dom";
import { ProductionTabs } from "@/features/production/components/ProductionTabs";
import { GlobalCalendarView } from "@/features/production/components/GlobalCalendarView";
import { CreatePlanWizard } from "@/features/production/components/CreatePlanWizard";
import { PlansList } from "@/features/production/components/PlansList";
import { WorkCalendarEditor } from "@/features/production/components/WorkCalendarEditor";
import type { ProductionTab } from "@/features/production/types/production";

const VALID_TABS: readonly ProductionTab[] = [
  "calendar",
  "create",
  "plans",
  "work-calendar",
];

const parseTab = (value: string | null): ProductionTab => {
  if (value && (VALID_TABS as readonly string[]).includes(value)) {
    return value as ProductionTab;
  }
  return "calendar";
};

export const ProductionPage = () => {
  const [searchParams, setSearchParams] = useSearchParams();
  const tab = parseTab(searchParams.get("tab"));

  useEffect(() => {
    if (!searchParams.get("tab")) {
      setSearchParams({ tab: "calendar" }, { replace: true });
    }
  }, [searchParams, setSearchParams]);

  const onTabChange = (next: ProductionTab) => {
    setSearchParams({ tab: next });
  };

  return (
    <main style={{ maxWidth: 1280, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <header>
          <h1 style={{ margin: 0, fontSize: "1.75rem" }}>Планирование производства плит</h1>
          <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
            Сводный календарь загрузки, создание и управление производственными планами.
          </p>
        </header>

        <ProductionTabs value={tab} onChange={onTabChange} />

        {tab === "calendar" && <GlobalCalendarView />}
        {tab === "create" && (
          <CreatePlanWizard onCreated={() => setSearchParams({ tab: "calendar" })} />
        )}
        {tab === "plans" && (
          <PlansList onOpenPlanCalendar={() => setSearchParams({ tab: "calendar" })} />
        )}
        {tab === "work-calendar" && <WorkCalendarEditor />}
      </div>
    </main>
  );
};
