import { useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { GsmRegistriesView } from "@/features/gsm/components/GsmRegistriesView";
import { GsmTabs } from "@/features/gsm/components/GsmTabs";
import { FleetOverviewView } from "@/features/gsm/components/FleetOverviewView";
import { TransactionsImportDialog } from "@/features/gsm/components/TransactionsImportDialog";
import { TransactionsJournalView } from "@/features/gsm/components/TransactionsJournalView";
import { summarizeImportReport } from "@/features/gsm/lib/importReport";
import type { GsmTab, TransactionImportReport } from "@/features/gsm/types/gsm";

export const GsmPage = () => {
  const [tab, setTab] = useState<GsmTab>("overview");
  const [importOpen, setImportOpen] = useState(false);
  const [lastImport, setLastImport] = useState<TransactionImportReport | null>(null);
  const lastImportSummary = lastImport ? summarizeImportReport(lastImport) : null;

  return (
    <main style={{ maxWidth: 1280, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <header>
          <h1 style={{ margin: 0, fontSize: "1.75rem" }}>ГСМ</h1>
          <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
            Путевые листы по топливным картам: обзор флота, транзакции и справочники.
          </p>
        </header>
        <GsmTabs value={tab} onChange={setTab} />
        {tab === "overview" && <FleetOverviewView />}
        {tab === "transactions" && (
          <section style={{ display: "grid", gap: "1rem" }} aria-label="Транзакции ГСМ">
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                gap: "0.75rem",
                flexWrap: "wrap",
                alignItems: "center",
              }}
            >
              <p style={{ margin: 0, color: "#475467" }}>
                Журнал транзакций флота. Импорт выгрузок .xls остаётся здесь.
              </p>
              <Button type="button" onClick={() => setImportOpen(true)}>
                Импорт транзакций
              </Button>
            </div>
            {lastImportSummary && (
              <Alert tone={lastImportSummary.tone}>Последний импорт: {lastImportSummary.text}</Alert>
            )}
            <TransactionsJournalView onOpenCards={() => setTab("registries")} />
            <TransactionsImportDialog
              open={importOpen}
              onClose={() => setImportOpen(false)}
              onImported={(report) => {
                setLastImport(report);
              }}
            />
          </section>
        )}
        {tab === "registries" && <GsmRegistriesView />}
      </div>
    </main>
  );
};
