import { useEffect, useState } from "react";
import { useSearchParams } from "react-router";
import { Card } from "@/shared/ui/Card";
import { Alert } from "@/shared/ui/Alert";
import { Spinner } from "@/shared/ui/Spinner";
import { ArchiveSectionTabs } from "@/features/commercial-archive/components/ArchiveSectionTabs";
import { ArchiveProductTypeFilterTabs } from "@/features/commercial-archive/components/ArchiveProductTypeFilter";
import { ArchiveOfferList } from "@/features/commercial-archive/components/ArchiveOfferList";
import { ArchiveSearchBar } from "@/features/commercial-archive/components/ArchiveSearchBar";
import { ContractCardList } from "@/features/commercial-archive/components/ContractCardList";
import { OfferDetailsDrawer } from "@/features/commercial-archive/components/OfferDetailsDrawer";
import { CurrentPlanButton } from "@/features/commercial-archive/components/CurrentPlanButton";
import {
  useArchiveListQuery,
  useArchiveSearchQuery,
} from "@/features/commercial-archive/hooks/useArchiveQueries";
import type {
  ArchiveOfferListItem,
  ArchivePageSection,
  ArchiveProductTypeFilter,
  ArchiveSearchState,
  ArchiveSection,
  ProductType,
} from "@/features/commercial-archive/types/archive";
import { getErrorMessage } from "@/shared/lib/apiError";

const VALID_SECTIONS: readonly ArchivePageSection[] = [
  "archived",
  "on_approval",
  "in_production",
  "completed",
  "counterparties",
];

/** Concrete types for contains-type filter (never the literal "mixed" token). */
export const concreteProductTypes = (item: ArchiveOfferListItem): ProductType[] => {
  const fromList = (item.product_types ?? []).filter(
    (type): type is ProductType => Boolean(type) && type !== "mixed",
  );
  if (fromList.length > 0) {
    return fromList;
  }
  if (item.product_type && item.product_type !== "mixed") {
    return [item.product_type];
  }
  // Legacy mono without product_type defaults to plates; bare "mixed" has no concrete types.
  if (!item.product_type) {
    return ["plates"];
  }
  return [];
};

/** Contains-type match for client-side search results (list trusts API filter). */
export const filterByProductType = (
  items: ArchiveOfferListItem[],
  productTypeFilter: ArchiveProductTypeFilter,
): ArchiveOfferListItem[] => {
  if (productTypeFilter === "all") {
    return items;
  }
  return items.filter((item) => concreteProductTypes(item).includes(productTypeFilter));
};

const parseSection = (value: string | null): ArchivePageSection => {
  if (value && (VALID_SECTIONS as readonly string[]).includes(value)) {
    return value as ArchivePageSection;
  }
  return "archived";
};

export const sectionFromStatus = (item: ArchiveOfferListItem): ArchiveSection => {
  switch (item.status) {
    case "в архиве":
      return "archived";
    case "на согласовании":
      return "on_approval";
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

/** Список КП отдельно: на «Контрагенты» компонент не монтируется, запрос раздела не уходит. */
const ArchiveKpListPane = ({
  section,
  productTypeFilter,
  onSelect,
}: {
  section: ArchiveSection;
  productTypeFilter: ArchiveProductTypeFilter;
  onSelect: (kpId: number) => void;
}) => {
  const listQuery = useArchiveListQuery(section, productTypeFilter);
  const listItems = listQuery.data ?? [];

  return (
    <Card>
      {listQuery.isPending && (
        <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
          <Spinner /> Загружаю список...
        </div>
      )}
      {listQuery.isError && <Alert tone="error">{getErrorMessage(listQuery.error)}</Alert>}
      {!listQuery.isPending && !listQuery.isError && (
        <ArchiveOfferList section={section} items={listItems} onSelect={onSelect} />
      )}
    </Card>
  );
};

export const CommercialOfferArchivePage = () => {
  const [searchParams, setSearchParams] = useSearchParams();
  const section = parseSection(searchParams.get("section"));
  const isCounterparties = section === "counterparties";

  const [selectedKpId, setSelectedKpId] = useState<number | null>(null);
  const [searchState, setSearchState] = useState<ArchiveSearchState>(null);
  const [productTypeFilter, setProductTypeFilter] = useState<ArchiveProductTypeFilter>("all");

  const searchQuery = useArchiveSearchQuery(isCounterparties ? null : searchState);

  useEffect(() => {
    if (!searchParams.get("section")) {
      const next = new URLSearchParams(searchParams);
      next.set("section", "archived");
      setSearchParams(next, { replace: true });
    }
  }, [searchParams, setSearchParams]);

  useEffect(() => {
    const raw = searchParams.get("kp");
    if (!raw) return;
    const kpId = Number(raw);
    if (Number.isInteger(kpId) && kpId > 0) {
      setSelectedKpId(kpId);
    }
  }, [searchParams]);

  const onSectionChange = (next: ArchivePageSection) => {
    setSearchParams({ section: next });
    setSearchState(null);
  };

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
              {section !== "counterparties" && (
                <>
                  <ArchiveProductTypeFilterTabs value={productTypeFilter} onChange={setProductTypeFilter} />
                  <ArchiveSearchBar
                    activeQuery={searchState}
                    onSubmit={setSearchState}
                    onClear={() => setSearchState(null)}
                  />
                </>
              )}
            </div>
            <CurrentPlanButton />
          </div>
        </Card>

        {section !== "counterparties" && searchState !== null && (
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

        {section === "counterparties" ? (
          <ContractCardList />
        ) : (
          <ArchiveKpListPane
            section={section}
            productTypeFilter={productTypeFilter}
            onSelect={(kpId) => setSelectedKpId(kpId)}
          />
        )}
      </div>

      <OfferDetailsDrawer
        open={selectedKpId !== null}
        kpId={selectedKpId}
        onClose={() => setSelectedKpId(null)}
      />
    </main>
  );
};
