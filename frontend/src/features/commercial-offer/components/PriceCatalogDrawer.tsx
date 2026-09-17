import { Drawer } from "@/shared/ui/Drawer";
import { formatOfferNumber } from "@/features/commercial-offer/lib/formatOfferNumbers";
import { catalogRowSharesUnpricedStem } from "@/features/commercial-offer/lib/commonMarkSearchPrefix";
import type { PriceCatalogItem } from "@/features/commercial-offer/types/commercialOffer";

type PriceCatalogDrawerProps = {
  open: boolean;
  onClose: () => void;
  q: string;
  onQueryChange: (value: string) => void;
  items: PriceCatalogItem[];
  missingMarks: string[];
  loading?: boolean;
  errorMessage?: string | null;
};

export const PriceCatalogDrawer = ({
  open,
  onClose,
  q,
  onQueryChange,
  items,
  missingMarks,
  loading = false,
  errorMessage = null,
}: PriceCatalogDrawerProps) => {
  const uniqueMissing = [...new Set(missingMarks.map((mark) => mark.trim()).filter(Boolean))];

  return (
    <Drawer open={open} onClose={onClose} title="Прайс" side="left" width={480}>
      <div style={{ display: "grid", gap: "0.85rem" }}>
        <label style={{ display: "grid", gap: "0.35rem" }}>
          <span style={{ fontWeight: 600 }}>Поиск</span>
          <input
            type="search"
            value={q}
            onChange={(event) => onQueryChange(event.target.value)}
            placeholder="Марка"
            aria-label="Поиск по прайсу"
            style={{
              width: "100%",
              border: "1px solid #d0d5dd",
              borderRadius: 10,
              padding: "0.55rem 0.75rem",
            }}
          />
        </label>

        {uniqueMissing.length > 0 ? (
          <div
            style={{
              border: "1px solid #fecdca",
              background: "#fef3f2",
              borderRadius: 12,
              padding: "0.75rem",
            }}
          >
            <div style={{ fontWeight: 600, color: "#b42318", marginBottom: "0.4rem" }}>
              В этом КП нет в прайсе
            </div>
            <ul style={{ margin: 0, paddingLeft: "1.2rem", color: "#b42318" }}>
              {uniqueMissing.map((mark) => (
                <li key={mark}>{mark}</li>
              ))}
            </ul>
          </div>
        ) : null}

        {errorMessage ? <div style={{ color: "#b42318" }}>{errorMessage}</div> : null}
        {loading ? <div style={{ color: "#667085" }}>Загружаем прайс…</div> : null}

        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.92rem" }}>
            <thead>
              <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                <th style={{ padding: "0.5rem 0.6rem" }}>Марка</th>
                <th style={{ padding: "0.5rem 0.6rem" }}>Класс</th>
                <th style={{ padding: "0.5rem 0.6rem" }}>Цена</th>
              </tr>
            </thead>
            <tbody>
              {items.map((item, index) => {
                const highlighted = catalogRowSharesUnpricedStem(item.mark, uniqueMissing);
                return (
                  <tr
                    key={`${item.mark}-${item.concrete_grade ?? "none"}-${index}`}
                    style={{ background: highlighted ? "#fff7ed" : undefined }}
                  >
                    <td style={{ padding: "0.5rem 0.6rem", borderBottom: "1px solid #f2f4f7" }}>{item.mark}</td>
                    <td style={{ padding: "0.5rem 0.6rem", borderBottom: "1px solid #f2f4f7" }}>
                      {item.concrete_grade ?? "—"}
                    </td>
                    <td style={{ padding: "0.5rem 0.6rem", borderBottom: "1px solid #f2f4f7" }}>
                      {formatOfferNumber(item.price)}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
          {!loading && items.length === 0 ? (
            <div style={{ color: "#667085", marginTop: "0.5rem" }}>Нет совпадений в прайсе.</div>
          ) : null}
        </div>
      </div>
    </Drawer>
  );
};
