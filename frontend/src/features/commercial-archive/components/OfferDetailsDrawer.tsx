import { useState } from "react";
import { Modal } from "@/shared/ui/Modal";
import { Button } from "@/shared/ui/Button";
import { Spinner } from "@/shared/ui/Spinner";
import { Alert } from "@/shared/ui/Alert";
import { archiveApi } from "@/features/commercial-archive/api/archiveApi";
import { useArchiveOfferQuery } from "@/features/commercial-archive/hooks/useArchiveQueries";
import { formatMoney, statusEmoji } from "@/features/commercial-archive/lib/format";
import { downloadFile } from "@/shared/lib/downloadFile";
import { getErrorMessage } from "@/shared/lib/apiError";
import { DiscountEditDialog } from "./DiscountEditDialog";
import { DeleteConfirmDialog } from "./DeleteConfirmDialog";
import { MoveToProductionDialog } from "./MoveToProductionDialog";

type Props = {
  open: boolean;
  kpId: number | null;
  onClose: () => void;
};

const PLATES_PREVIEW = 10;

export const OfferDetailsDrawer = ({ open, kpId, onClose }: Props) => {
  const query = useArchiveOfferQuery(open ? kpId : null);
  const [showAllPlates, setShowAllPlates] = useState(false);
  const [discountOpen, setDiscountOpen] = useState(false);
  const [deleteOpen, setDeleteOpen] = useState(false);
  const [moveOpen, setMoveOpen] = useState(false);

  const offer = query.data;
  const platesToShow = offer
    ? showAllPlates
      ? offer.plates
      : offer.plates.slice(0, PLATES_PREVIEW)
    : [];

  return (
    <Modal open={open} onClose={onClose} title={offer ? `КП №${offer.kp_id}` : "Карточка КП"} maxWidth={720}>
      {query.isPending && (
        <div style={{ display: "flex", gap: "0.75rem", alignItems: "center" }}>
          <Spinner /> Загружаю данные...
        </div>
      )}

      {query.isError && <Alert tone="error">{getErrorMessage(query.error)}</Alert>}

      {offer && (
        <div style={{ display: "grid", gap: "1rem" }}>
          <section
            style={{
              display: "grid",
              gap: "0.45rem",
              padding: "1rem",
              background: "#f8faff",
              border: "1px solid #e4e7ec",
              borderRadius: 14,
            }}
          >
            <div style={{ display: "flex", flexWrap: "wrap", gap: "0.75rem", justifyContent: "space-between" }}>
              <div>
                <div style={{ color: "#667085", fontSize: "0.85rem" }}>Клиент</div>
                <div style={{ fontWeight: 600 }}>{offer.customer_name || "—"}</div>
              </div>
              <div>
                <div style={{ color: "#667085", fontSize: "0.85rem" }}>Менеджер</div>
                <div style={{ fontWeight: 600 }}>{offer.manager_name || "—"}</div>
              </div>
              <div>
                <div style={{ color: "#667085", fontSize: "0.85rem" }}>Дата создания</div>
                <div style={{ fontWeight: 600 }}>{offer.creation_date || "—"}</div>
              </div>
              <div>
                <div style={{ color: "#667085", fontSize: "0.85rem" }}>Статус</div>
                <div style={{ fontWeight: 600 }}>
                  {statusEmoji(offer.status)} {offer.status || "—"}
                </div>
              </div>
              {offer.execution_terms && (
                <div>
                  <div style={{ color: "#667085", fontSize: "0.85rem" }}>Срок</div>
                  <div style={{ fontWeight: 600 }}>⏰ {offer.execution_terms}</div>
                </div>
              )}
              {offer.completion_percentage !== null && (
                <div>
                  <div style={{ color: "#667085", fontSize: "0.85rem" }}>Готовность</div>
                  <div style={{ fontWeight: 600 }}>{offer.completion_percentage.toFixed(1)}%</div>
                </div>
              )}
            </div>
          </section>

          <section
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
              gap: "0.75rem",
            }}
          >
            <FinanceCard label="Сумма без НДС" value={formatMoney(offer.finance.subtotal)} />
            <FinanceCard label="НДС (22%)" value={formatMoney(offer.finance.vat_amount)} />
            <FinanceCard label="Итого с НДС" value={formatMoney(offer.finance.total_amount)} accent />
            <FinanceCard label="Скидка" value={`${offer.finance.discount_percent}%`} />
          </section>

          <section>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "0.5rem" }}>
              <h3 style={{ margin: 0 }}>Состав заказа ({offer.plates.length})</h3>
              {offer.plates.length > PLATES_PREVIEW && (
                <Button variant="ghost" onClick={() => setShowAllPlates((prev) => !prev)}>
                  {showAllPlates ? "Свернуть" : "Показать все"}
                </Button>
              )}
            </div>
            {offer.plates.length === 0 ? (
              <div style={{ color: "#667085" }}>Список плит пуст.</div>
            ) : (
              <div style={{ overflowX: "auto" }}>
                <table style={{ borderCollapse: "collapse", width: "100%", fontSize: "0.95rem" }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: "#475467", background: "#f2f4f7" }}>
                      <th style={{ padding: "0.5rem 0.75rem" }}>№</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Наименование</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Кол-во</th>
                      <th style={{ padding: "0.5rem 0.75rem" }}>Цена</th>
                    </tr>
                  </thead>
                  <tbody>
                    {platesToShow.map((plate, index) => (
                      <tr key={`${plate.plate_name}-${index}`} style={{ borderTop: "1px solid #e4e7ec" }}>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{plate.position_number ?? index + 1}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{plate.plate_name || "—"}</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>{plate.qty} шт</td>
                        <td style={{ padding: "0.5rem 0.75rem" }}>
                          {plate.discounted_price !== null
                            ? formatMoney(plate.discounted_price)
                            : plate.unit_price !== null
                              ? formatMoney(plate.unit_price)
                              : "—"}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>

          <section style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
            <Button onClick={() => downloadFile(archiveApi.buildDocumentUrl(offer.kp_id, "pdf"))}>
              📄 PDF
            </Button>
            <Button onClick={() => downloadFile(archiveApi.buildDocumentUrl(offer.kp_id, "xlsx"))}>
              📊 XLSX
            </Button>
            <Button variant="secondary" onClick={() => setDiscountOpen(true)}>
              Изменить скидку
            </Button>
            {offer.status === "в архиве" && (
              <Button variant="secondary" onClick={() => setMoveOpen(true)}>
                🏭 В производство
              </Button>
            )}
            <Button variant="danger" onClick={() => setDeleteOpen(true)}>
              Удалить КП
            </Button>
          </section>
        </div>
      )}

      {offer && (
        <>
          <DiscountEditDialog
            open={discountOpen}
            onClose={() => setDiscountOpen(false)}
            kpId={offer.kp_id}
            currentDiscount={offer.finance.discount_percent}
          />
          <DeleteConfirmDialog
            open={deleteOpen}
            onClose={() => setDeleteOpen(false)}
            onDeleted={onClose}
            kpId={offer.kp_id}
            customerName={offer.customer_name}
          />
          <MoveToProductionDialog
            open={moveOpen}
            onClose={() => setMoveOpen(false)}
            kpId={offer.kp_id}
          />
        </>
      )}
    </Modal>
  );
};

const FinanceCard = ({ label, value, accent }: { label: string; value: string; accent?: boolean }) => (
  <div
    style={{
      padding: "0.75rem 1rem",
      borderRadius: 14,
      border: "1px solid #e4e7ec",
      background: accent ? "#eef4ff" : "#ffffff",
    }}
  >
    <div style={{ fontSize: "0.85rem", color: "#667085" }}>{label}</div>
    <div style={{ fontWeight: 700, color: accent ? "#1d4ed8" : "#101828", marginTop: "0.25rem" }}>{value}</div>
  </div>
);
