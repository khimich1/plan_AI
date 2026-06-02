import { Alert } from "@/shared/ui/Alert";

type ProductionEstimateAlertProps = {
  estimatedTracks: number;
  estimatedDays: number;
  totalLengthM: number;
  label?: string;
  context?: string;
};

export const ProductionEstimateAlert = ({
  estimatedTracks,
  estimatedDays,
  totalLengthM,
  label = "Оценка производства",
  context,
}: ProductionEstimateAlertProps) => (
  <Alert tone="info">
    {label}: ~{estimatedTracks} дорожек, ~{estimatedDays} дней (суммарная длина{" "}
    {totalLengthM.toFixed(1)} м{context ? `, ${context}` : ""}).
  </Alert>
);
