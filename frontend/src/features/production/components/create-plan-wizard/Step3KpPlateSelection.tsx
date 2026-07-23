import { Alert } from "@/shared/ui/Alert";
import { Button } from "@/shared/ui/Button";
import { FieldWrapper } from "@/shared/ui/Field";
import { Spinner } from "@/shared/ui/Spinner";
import type { ProductionEstimate } from "@/features/production/lib/productionEstimate";
import type {
  FilterMethod,
  KpCandidateItem,
} from "@/features/production/types/production";
import { KpCandidatesTable } from "./KpCandidatesTable";
import { ProductionEstimateAlert } from "./ProductionEstimateAlert";

type Props = {
  isFillMode: boolean;
  filterMethod: FilterMethod;
  candidatesLoading: boolean;
  candidatesError: boolean;
  candidates: KpCandidateItem[] | undefined;
  selectionEstimate: ProductionEstimate | null;
  tracksPerDay: number;
  tracksPerDaySource: "календарь";
  estimateByKpId: Map<number, ProductionEstimate | null>;
  selectedPlatesByKp: Record<number, number[]>;
  selectedPlateQtyByKp: Record<number, Record<number, number>>;
  expandedKpIds: Set<number>;
  buildPending: boolean;
  buildSuccess: boolean;
  buildErrorMessage: string | null;
  canSubmit: boolean;
  onFilterMethodChange: (method: FilterMethod) => void;
  onToggleKp: (kp: KpCandidateItem) => void;
  onToggleExpand: (kpId: number) => void;
  onTogglePlate: (kp: KpCandidateItem, plateId: number) => void;
  onSetPlateQty: (kp: KpCandidateItem, plateId: number, qty: number) => void;
  onCancelFill: () => void;
  onSubmit: (order: "asc" | "desc") => void;
};

export const Step3KpPlateSelection = ({
  isFillMode,
  filterMethod,
  candidatesLoading,
  candidatesError,
  candidates,
  selectionEstimate,
  tracksPerDay,
  tracksPerDaySource,
  estimateByKpId,
  selectedPlatesByKp,
  selectedPlateQtyByKp,
  expandedKpIds,
  buildPending,
  buildSuccess,
  buildErrorMessage,
  canSubmit,
  onFilterMethodChange,
  onToggleKp,
  onToggleExpand,
  onTogglePlate,
  onSetPlateQty,
  onCancelFill,
  onSubmit,
}: Props) => (
  <div style={{ display: "grid", gap: "1rem" }}>
    <FieldWrapper label="Какие КП включить в план">
      <div style={{ display: "flex", gap: "0.5rem", flexWrap: "wrap" }}>
        <Button
          variant={filterMethod === "all" ? "primary" : "secondary"}
          onClick={() => onFilterMethodChange("all")}
        >
          Все КП в работе
        </Button>
        <Button
          variant={filterMethod === "kp" ? "primary" : "secondary"}
          onClick={() => onFilterMethodChange("kp")}
        >
          Выбрать КП вручную
        </Button>
      </div>
    </FieldWrapper>

    {candidatesLoading && (
      <div style={{ display: "flex", gap: "0.5rem", alignItems: "center" }}>
        <Spinner /> Загрузка данных для оценки…
      </div>
    )}

    {selectionEstimate && (
      <ProductionEstimateAlert
        estimate={selectionEstimate}
        tracksPerDay={tracksPerDay}
        tracksPerDaySource={tracksPerDaySource}
        label={
          filterMethod === "all"
            ? "Оценка по всем КП в работе"
            : "Оценка выбранного"
        }
      />
    )}

    {filterMethod === "kp" && (
      <KpCandidatesTable
        loading={candidatesLoading}
        error={candidatesError}
        candidates={candidates}
        estimateByKpId={estimateByKpId}
        selectedPlatesByKp={selectedPlatesByKp}
        selectedPlateQtyByKp={selectedPlateQtyByKp}
        expandedKpIds={expandedKpIds}
        onToggleKp={onToggleKp}
        onToggleExpand={onToggleExpand}
        onTogglePlate={onTogglePlate}
        onSetPlateQty={onSetPlateQty}
      />
    )}

    {buildErrorMessage && (
      <Alert tone="error">{buildErrorMessage}</Alert>
    )}
    {buildSuccess && (
      <Alert tone="success">
        План успешно создан. Переключаюсь на календарный план…
      </Alert>
    )}

    <div style={{ display: "flex", flexDirection: "column", gap: "0.5rem" }}>
      <p style={{ margin: 0, fontSize: "0.85rem", color: "#667085" }}>
        Режим «сильные первыми»: сильные группы и первая плита первыми; перед каждой
        группой резов подбирается целая с ближайшей нагрузкой. Экспериментальный режим —
        возможен рост переармирования ранних дорожек.
      </p>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <Button variant="ghost" onClick={onCancelFill}>
          ← К календарю
        </Button>
        <div style={{ display: "flex", gap: "0.5rem" }}>
          <Button
            variant="secondary"
            onClick={() => onSubmit("desc")}
            disabled={!canSubmit}
          >
            {buildPending
              ? "Строим план…"
              : isFillMode
                ? "Дозаполнение: сильные первыми"
                : "Планирование: сильные первыми"}
          </Button>
          <Button onClick={() => onSubmit("asc")} disabled={!canSubmit}>
            {buildPending
              ? "Строим план…"
              : isFillMode
                ? "Запустить дозаполнение"
                : "Запустить планирование"}
          </Button>
        </div>
      </div>
    </div>
  </div>
);
