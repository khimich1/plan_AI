import { useState } from "react";
import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { GsmPeriodView } from "@/features/gsm/components/GsmPeriodView";
import { GsmRegistriesView } from "@/features/gsm/components/GsmRegistriesView";
import { GsmTabs } from "@/features/gsm/components/GsmTabs";
import { TransactionsImportDialog } from "@/features/gsm/components/TransactionsImportDialog";
import type { GsmTab, TransactionImportReport } from "@/features/gsm/types/gsm";

export const GsmPage = () => {
  const [tab, setTab] = useState<GsmTab>("period");
  const [importOpen, setImportOpen] = useState(false);
  const [lastImport, setLastImport] = useState<TransactionImportReport | null>(null);

  return (
    <main style={{ maxWidth: 1280, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <header>
          <h1 style={{ margin: 0, fontSize: "1.75rem" }}>ГСМ</h1>
          <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
            Путевые листы по топливным картам: период, транзакции и справочники.
          </p>
        </header>
        <GsmTabs value={tab} onChange={setTab} />
        {tab === "period" && <GsmPeriodView />}
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
                Журнал транзакций появится позже. Сейчас доступен мульти-файл импорт выгрузок .xls.
              </p>
              <Button type="button" onClick={() => setImportOpen(true)}>
                Импорт транзакций
              </Button>
            </div>
            {lastImport && (
              <Alert tone="success">
                Последний импорт: вставлено {lastImport.rows_inserted}, дублей{" "}
                {lastImport.rows_duplicate}, файлов {lastImport.files.length}.
              </Alert>
            )}
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
