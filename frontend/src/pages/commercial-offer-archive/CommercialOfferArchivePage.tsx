import { useEffect, useMemo, useState } from "react";
import { useSearchParams } from "react-router-dom";
import { Card } from "@/shared/ui/Card";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import { ArchiveSectionTabs } from "@/features/commercial-archive/components/ArchiveSectionTabs";
import { ArchiveOfferList } from "@/features/commercial-archive/components/ArchiveOfferList";
import { ArchiveSearchBar } from "@/features/commercial-archive/components/ArchiveSearchBar";
import { OfferDetailsDrawer } from "@/features/commercial-archive/components/OfferDetailsDrawer";
import { CurrentPlanButton } from "@/features/commercial-archive/components/CurrentPlanButton";
import {
  useArchiveListQuery,
  useArchiveSearchQuery,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import type { ArchiveSection } from "@/features/commercial-archive/types/archive";
import { getErrorMessage } from "@/shared/lib/apiError";

const VALID_SECTIONS: readonly ArchiveSection[] = ["archived", "in_production", "completed"];

const parseSection = (value: string | null): ArchiveSection => {
  if (value && (VALID_SECTIONS as readonly string[]).includes(value)) {
    return value as ArchiveSection;
  }
  return "archived";
};

export const CommercialOfferArchivePage = () => {
  const [searchParams, setSearchParams] = useSearchParams();
  const section = parseSection(searchParams.get("section"));

  const [selectedKpId, setSelectedKpId] = useState<number | null>(null);
  const [searchKpId, setSearchKpId] = useState<number | null>(null);

  const listQuery = useArchiveListQuery(section);
  const searchQuery = useArchiveSearchQuery(searchKpId);

  useEffect(() => {
    if (!searchParams.get("section")) {
      setSearchParams({ section: "archived" }, { replace: true });
    }
  }, [searchParams, setSearchParams]);

  const onSectionChange = (next: ArchiveSection) => {
    setSearchParams({ section: next });
    setSearchKpId(null);
  };

  const filteredItems = useMemo(() => listQuery.data ?? [], [listQuery.data]);

  return (
    <main style={{ maxWidth: 1280, margin: "0 auto", padding: "2rem 1rem 4rem" }}>
      <div style={{ display: "grid", gap: "1rem" }}>
        <header>
          <h1 style={{ margin: 0, fontSize: "1.75rem" }}>Архив коммерческих предложений</h1>
          <p style={{ margin: "0.4rem 0 0", color: "#475467" }}>
            Просматривайте КП по разделам, скачивайте документы и управляйте статусами.
          </p>
        </header>

        <Card>
          <div
            style={{
              display: "flex",
              gap: "1rem",
              justifyContent: "space-between",
              alignItems: "flex-start",
              flexWrap: "wrap",
            }}
          >
            <div style={{ display: "grid", gap: "0.75rem", flex: "1 0 320px" }}>
              <ArchiveSectionTabs value={section} onChange={onSectionChange} />
              <ArchiveSearchBar
                activeQuery={searchKpId}
                onSubmit={setSearchKpId}
                onClear={() => setSearchKpId(null)}
              />
            </div>
            <CurrentPlanButton />
          </div>
        </Card>

        {searchKpId !== null && (
          <Card title={`Поиск по номеру: КП №${searchKpId}`}>
            {searchQuery.isPending && (
              <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
                <Spinner /> Ищу КП...
              </div>
            )}
            {searchQuery.isError && <Alert tone="error">{getErrorMessage(searchQuery.error)}</Alert>}
            {searchQuery.data && !searchQuery.data.found && (
              <Alert tone="warning">КП с номером {searchKpId} не найдено.</Alert>
            )}
            {searchQuery.data?.found && searchQuery.data.offer && (
              <div style={{ display: "grid", gap: "0.5rem" }}>
                <div>
                  <strong>Клиент:</strong> {searchQuery.data.offer.customer_name || "—"}
                </div>
                <div>
                  <strong>Статус:</strong> {searchQuery.data.offer.status || "—"}
                </div>
                <button
                  type="button"
                  onClick={() => setSelectedKpId(searchQuery.data!.offer!.kp_id)}
                  style={{
                    marginTop: "0.25rem",
                    padding: "0.6rem 0.9rem",
                    borderRadius: 12,
                    border: "1px solid #2b5cff",
                    background: "#2b5cff",
                    color: "#ffffff",
                    fontWeight: 600,
                    cursor: "pointer",
                    justifySelf: "start",
                  }}
                >
                  Открыть карточку
                </button>
              </div>
            )}
          </Card>
        )}

        <Card>
          {listQuery.isPending && (
            <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
              <Spinner /> Загружаю список...
            </div>
          )}
          {listQuery.isError && <Alert tone="error">{getErrorMessage(listQuery.error)}</Alert>}
          {!listQuery.isPending && !listQuery.isError && (
            <ArchiveOfferList
              section={section}
              items={filteredItems}
              onSelect={(kpId) => setSelectedKpId(kpId)}
            />
          )}
        </Card>
      </div>

      <OfferDetailsDrawer
        open={selectedKpId !== null}
        kpId={selectedKpId}
        onClose={() => setSelectedKpId(null)}
      />
    </main>
  );
};
