import { useEffect, useMemo, useState } from "react";
import { useSearchParams } from "react-router";
import { Card } from "@/shared/ui/Card";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import { ArchiveSectionTabs } from "@/features/commercial-archive/components/ArchiveSectionTabs";
import { ArchiveProductTypeFilterTabs } from "@/features/commercial-archive/components/ArchiveProductTypeFilter";
import { ArchiveOfferList } from "@/features/commercial-archive/components/ArchiveOfferList";
import { ArchiveSearchBar } from "@/features/commercial-archive/components/ArchiveSearchBar";
import { OfferDetailsDrawer } from "@/features/commercial-archive/components/OfferDetailsDrawer";
import { CurrentPlanButton } from "@/features/commercial-archive/components/CurrentPlanButton";
import {
  useArchiveListQuery,
  useArchiveSearchQuery,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import type {
  ArchiveOfferListItem,
  ArchiveProductTypeFilter,
  ArchiveSearchState,
  ArchiveSection,
  ProductType,
} from "@/features/commercial-archive/types/archive";
import { getErrorMessage } from "@/shared/lib/apiError";

const VALID_SECTIONS: readonly ArchiveSection[] = ["archived", "in_production", "completed"];

const resolveProductType = (item: ArchiveOfferListItem): ProductType => item.product_type ?? "plates";

const filterByProductType = (
  items: ArchiveOfferListItem[],
  productTypeFilter: ArchiveProductTypeFilter,
): ArchiveOfferListItem[] => {
  if (productTypeFilter === "all") {
    return items;
  }
  return items.filter((item) => resolveProductType(item) === productTypeFilter);
};

const parseSection = (value: string | null): ArchiveSection => {
  if (value && (VALID_SECTIONS as readonly string[]).includes(value)) {
    return value as ArchiveSection;
  }
  return "archived";
};

export const sectionFromStatus = (item: ArchiveOfferListItem): ArchiveSection => {
  switch (item.status) {
    case "в архиве":
      return "archived";
    case "в работе":
      return "in_production";
    case "На СГП":
      return "in_production";
    case "выполнено":
      return "completed";
    default:
      return "archived";
  }
};

const searchCardTitle = (searchState: ArchiveSearchState): string => {
  if (!searchState) {
    return "Поиск";
  }
  if (searchState.kind === "number") {
    return `Поиск: КП №${searchState.value}`;
  }
  return `Поиск: «${searchState.value}»`;
};

export const CommercialOfferArchivePage = () => {
  const [searchParams, setSearchParams] = useSearchParams();
  const section = parseSection(searchParams.get("section"));

  const [selectedKpId, setSelectedKpId] = useState<number | null>(null);
  const [searchState, setSearchState] = useState<ArchiveSearchState>(null);
  const [productTypeFilter, setProductTypeFilter] = useState<ArchiveProductTypeFilter>("all");

  const listQuery = useArchiveListQuery(section, productTypeFilter);
  const searchQuery = useArchiveSearchQuery(searchState);

  useEffect(() => {
    if (!searchParams.get("section")) {
      setSearchParams({ section: "archived" }, { replace: true });
    }
  }, [searchParams, setSearchParams]);

  const onSectionChange = (next: ArchiveSection) => {
    setSearchParams({ section: next });
    setSearchState(null);
  };

  const filteredItems = useMemo(
    () => filterByProductType(listQuery.data ?? [], productTypeFilter),
    [listQuery.data, productTypeFilter],
  );

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
              <ArchiveProductTypeFilterTabs value={productTypeFilter} onChange={setProductTypeFilter} />
              <ArchiveSearchBar
                activeQuery={searchState}
                onSubmit={setSearchState}
                onClear={() => setSearchState(null)}
              />
            </div>
            <CurrentPlanButton />
          </div>
        </Card>

        {searchState !== null && (
          <Card title={searchCardTitle(searchState)}>
            {searchQuery.isPending && (
              <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
                <Spinner /> Ищу КП...
              </div>
            )}
            {searchQuery.isError && <Alert tone="error">{getErrorMessage(searchQuery.error)}</Alert>}
            {searchQuery.data && searchQuery.data.total === 0 && (
              <Alert tone="warning">Ничего не найдено</Alert>
            )}
            {searchQuery.data && searchQuery.data.total > 0 && (
              <div style={{ display: "grid", gap: "0.75rem" }}>
                {searchQuery.data.truncated && (
                  <Alert tone="warning">
                    Показаны первые 50 из {searchQuery.data.total}
                  </Alert>
                )}
                <ArchiveOfferList
                  section={section}
                  items={filterByProductType(
                    // Archive page: admin/manager only — full commercial items.
                    searchQuery.data.items as ArchiveOfferListItem[],
                    productTypeFilter,
                  )}
                  onSelect={(kpId) => setSelectedKpId(kpId)}
                  sectionForItem={sectionFromStatus}
                />
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
