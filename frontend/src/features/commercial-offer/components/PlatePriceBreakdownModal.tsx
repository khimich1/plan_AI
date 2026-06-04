import { Modal } from "@/shared/ui/Modal";
import { Alert } from "@/shared/ui/Alert";
import type { BreakdownTable } from "@/features/commercial-offer/types/commercialOffer";
import { isBreakdownTotalRow } from "@/features/commercial-offer/lib/findBreakdownTable";

type PlatePriceBreakdownModalProps = {
  open: boolean;
  plateName: string | null;
  table: BreakdownTable | undefined;
  onClose: () => void;
};

export const PlatePriceBreakdownModal = ({
  open,
  plateName,
  table,
  onClose,
}: PlatePriceBreakdownModalProps) => {
  if (!plateName) {
    return null;
  }

  return (
    <Modal open={open} onClose={onClose} title={plateName} maxWidth={520}>
      {!table ? (
        <Alert tone="warning">Детальная разбивка для этой позиции недоступна.</Alert>
      ) : (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.9rem" }}>
            <thead>
              <tr>
                {["Компонент", "Расчёт", "Сумма"].map((column) => (
                  <th
                    key={column}
                    style={{
                      textAlign: "left",
                      padding: "0.55rem 0.65rem",
                      borderBottom: "1px solid #e4e7ec",
                      color: "#475467",
                      fontWeight: 600,
                    }}
                  >
                    {column}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {table.rows.map((row, index) => {
                const [component = "", calculation = "", sum = ""] = row;
                const isTotal = isBreakdownTotalRow(component);
                return (
                  <tr
                    key={`${component}-${index}`}
                    style={{
                      background: isTotal ? "#f8fafc" : undefined,
                      fontWeight: isTotal ? 600 : undefined,
                    }}
                  >
                    <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7", verticalAlign: "top" }}>
                      {component}
                    </td>
                    <td style={{ padding: "0.55rem 0.65rem", borderBottom: "1px solid #f2f4f7", verticalAlign: "top" }}>
                      {calculation}
                    </td>
                    <td
                      style={{
                        padding: "0.55rem 0.65rem",
                        borderBottom: "1px solid #f2f4f7",
                        verticalAlign: "top",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {sum}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </Modal>
  );
};
