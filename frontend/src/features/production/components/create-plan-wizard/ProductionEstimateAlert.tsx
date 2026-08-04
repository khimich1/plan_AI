import { Alert } from "@/shared/ui/Alert";
import type { ProductionEstimate } from "@/features/production/lib/productionEstimate";

type Props = {
  estimate: ProductionEstimate;
  tracksPerDay: number;
  tracksPerDaySource: "календарь";
  label: string;
};

export const ProductionEstimateAlert = ({
  estimate,
  tracksPerDay,
  tracksPerDaySource,
  label,
}: Props) => (
  <Alert tone="info">
    {label}: ~{estimate.estimated_tracks} дорожек, ~{estimate.estimated_days} дней
    (суммарная длина {estimate.total_length_m.toFixed(1)} м, при {tracksPerDay} дор./день —{" "}
    {tracksPerDaySource}).
  </Alert>
);
