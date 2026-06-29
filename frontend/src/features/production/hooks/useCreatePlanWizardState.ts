import { useEffect, useMemo, useState } from "react";
import { getErrorMessage } from "@/shared/lib/apiError";
import { isPlanVersionConflict } from "@/shared/lib/planConflict";
import {
  useBuildPlanMutation,
  useDayOccupancyQuery,
  useGlobalCalendarQuery,
  useKpCandidatesQuery,
  usePlansListQuery,
  useWorkCalendarQuery,
} from "@/features/production/hooks/useProductionQueries";
import {
  allPlatesLengthM,
  estimateFromLengthM,
  estimateKpSelection,
  selectedLengthM,
  type ProductionEstimate,
} from "@/features/production/lib/productionEstimate";
import type {
  FillTargetItem,
  FilterMethod,
  KpCandidateItem,
} from "@/features/production/types/production";
import { formatRu, MAX_PER_DAY, startOfMonth, todayISO } from "@/features/production/components/create-plan-wizard/utils";

export type UseCreatePlanWizardStateOptions = {
  onCreated?: () => void;
  fillRequest?: FillTargetItem[] | null;
  onFillRequestConsumed?: () => void;
  onCancelFill?: () => void;
};

export const useCreatePlanWizardState = ({
  onCreated,
  fillRequest,
  onFillRequestConsumed,
  onCancelFill,
}: UseCreatePlanWizardStateOptions) => {
  const [step, setStep] = useState<1 | 2 | 3>(1);
  const [startDate, setStartDate] = useState<string>(todayISO());
  const [tracksCount, setTracksCount] = useState<number>(1);
  const [filterMethod, setFilterMethod] = useState<FilterMethod>("all");
  const [selectedPlatesByKp, setSelectedPlatesByKp] = useState<
    Record<number, number[]>
  >({});
  const [selectedPlateQtyByKp, setSelectedPlateQtyByKp] = useState<
    Record<number, Record<number, number>>
  >({});
  const [expandedKpIds, setExpandedKpIds] = useState<Set<number>>(new Set());
  const [planName, setPlanName] = useState<string>("");
  const [calendarMonth, setCalendarMonth] = useState<Date>(() =>
    startOfMonth(new Date()),
  );
  const [fillTargets, setFillTargets] = useState<FillTargetItem[] | null>(null);

  const occupancyQuery = useDayOccupancyQuery();
  const calendarQuery = useGlobalCalendarQuery();
  const plansQuery = usePlansListQuery();
  const workCalendar = useWorkCalendarQuery();
  const candidatesQuery = useKpCandidatesQuery(step === 3);
  const buildMutation = useBuildPlanMutation();

  const occupancy = occupancyQuery.data?.occupancy ?? {};
  const maxPerDay = occupancyQuery.data?.max_per_day ?? MAX_PER_DAY;
  const daysInfo = calendarQuery.data?.days_info ?? {};
  const holidays = useMemo(
    () => new Set(workCalendar.data?.extra_holidays ?? []),
    [workCalendar.data],
  );
  const extraWorkdays = useMemo(
    () => new Set(workCalendar.data?.extra_workdays ?? []),
    [workCalendar.data],
  );

  const occupiedOnStart = occupancy[startDate] ?? 0;
  const freeOnStart = Math.max(0, maxPerDay - occupiedOnStart);

  useEffect(() => {
    if (fillRequest && fillRequest.length > 0) {
      setFillTargets(fillRequest);
      setStep(3);
      onFillRequestConsumed?.();
    }
  }, [fillRequest, onFillRequestConsumed]);

  const defaultQtyMap = (kp: KpCandidateItem, plateIds: number[]) => {
    const map: Record<number, number> = {};
    for (const plate of kp.plates) {
      if (plateIds.includes(plate.id)) {
        map[plate.id] = plate.qty;
      }
    }
    return map;
  };

  const toggleKp = (kp: KpCandidateItem) => {
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      if (kp.kp_id in next) {
        delete next[kp.kp_id];
        setSelectedPlateQtyByKp((q) => {
          const qn = { ...q };
          delete qn[kp.kp_id];
          return qn;
        });
      } else {
        const ids = kp.plates.map((p) => p.id);
        next[kp.kp_id] = ids;
        setSelectedPlateQtyByKp((q) => ({
          ...q,
          [kp.kp_id]: defaultQtyMap(kp, ids),
        }));
      }
      return next;
    });
  };

  const togglePlate = (kp: KpCandidateItem, plateId: number) => {
    const plate = kp.plates.find((p) => p.id === plateId);
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      const current = next[kp.kp_id];
      if (current === undefined) {
        next[kp.kp_id] = [plateId];
        if (plate) {
          setSelectedPlateQtyByKp((q) => ({
            ...q,
            [kp.kp_id]: { ...(q[kp.kp_id] ?? {}), [plateId]: plate.qty },
          }));
        }
      } else if (current.includes(plateId)) {
        const filtered = current.filter((id) => id !== plateId);
        if (filtered.length === 0) {
          delete next[kp.kp_id];
          setSelectedPlateQtyByKp((q) => {
            const qn = { ...q };
            delete qn[kp.kp_id];
            return qn;
          });
        } else {
          next[kp.kp_id] = filtered;
          setSelectedPlateQtyByKp((q) => {
            const perKp = { ...(q[kp.kp_id] ?? {}) };
            delete perKp[plateId];
            return { ...q, [kp.kp_id]: perKp };
          });
        }
      } else {
        next[kp.kp_id] = [...current, plateId];
        if (plate) {
          setSelectedPlateQtyByKp((q) => ({
            ...q,
            [kp.kp_id]: { ...(q[kp.kp_id] ?? {}), [plateId]: plate.qty },
          }));
        }
      }
      return next;
    });
  };

  const setPlateQty = (kp: KpCandidateItem, plateId: number, rawQty: number) => {
    const plate = kp.plates.find((p) => p.id === plateId);
    if (!plate) {
      return;
    }
    const clamped = Math.max(1, Math.min(plate.qty, Math.round(rawQty) || 1));
    setSelectedPlateQtyByKp((prev) => ({
      ...prev,
      [kp.kp_id]: {
        ...(prev[kp.kp_id] ?? {}),
        [plateId]: clamped,
      },
    }));
  };

  const toggleExpand = (kpId: number) => {
    setExpandedKpIds((prev) => {
      const next = new Set(prev);
      if (next.has(kpId)) {
        next.delete(kpId);
      } else {
        next.add(kpId);
      }
      return next;
    });
  };

  const isFillMode = fillTargets !== null;

  const tracksPerDay =
    isFillMode && fillTargets && fillTargets.length > 0
      ? Math.max(...fillTargets.map((t) => t.tracks))
      : tracksCount;
  const tracksPerDaySource: "шаг 2" | "дозаполнение" =
    isFillMode && fillTargets && fillTargets.length > 0
      ? "дозаполнение"
      : "шаг 2";

  const selectionEstimate = useMemo((): ProductionEstimate | null => {
    const items = candidatesQuery.data?.items;
    if (!items?.length) {
      return null;
    }
    if (filterMethod === "all") {
      const length = items.reduce((sum, kp) => sum + allPlatesLengthM(kp), 0);
      if (length <= 0) {
        return null;
      }
      return estimateFromLengthM(length, tracksPerDay);
    }
    let totalLength = 0;
    for (const kp of items) {
      const ids = selectedPlatesByKp[kp.kp_id];
      if (!ids?.length) {
        continue;
      }
      totalLength += selectedLengthM(
        kp,
        ids,
        selectedPlateQtyByKp[kp.kp_id] ?? {},
      );
    }
    if (totalLength <= 0) {
      return null;
    }
    return estimateFromLengthM(totalLength, tracksPerDay);
  }, [
    candidatesQuery.data,
    filterMethod,
    selectedPlatesByKp,
    selectedPlateQtyByKp,
    tracksPerDay,
  ]);

  const estimateByKpId = useMemo(() => {
    const map = new Map<number, ProductionEstimate | null>();
    const items = candidatesQuery.data?.items;
    if (!items) {
      return map;
    }
    for (const kp of items) {
      const ids = selectedPlatesByKp[kp.kp_id] ?? [];
      map.set(
        kp.kp_id,
        estimateKpSelection(
          kp,
          ids,
          selectedPlateQtyByKp[kp.kp_id] ?? {},
          tracksPerDay,
        ),
      );
    }
    return map;
  }, [
    candidatesQuery.data,
    selectedPlatesByKp,
    selectedPlateQtyByKp,
    tracksPerDay,
  ]);

  const canProceedStep1 = Boolean(startDate);
  const canProceedStep2 = tracksCount >= 1 && tracksCount <= 50;
  const hasAnyPlateSelected =
    filterMethod === "all" ||
    Object.values(selectedPlatesByKp).some((ids) => ids.length > 0);
  const canSubmit =
    (isFillMode || (canProceedStep1 && canProceedStep2)) &&
    hasAnyPlateSelected &&
    !buildMutation.isPending &&
    !buildMutation.isSuccess;

  const fillTotalTracks = fillTargets
    ? fillTargets.reduce((acc, t) => acc + t.tracks, 0)
    : 0;
  const fillSubtitle = fillTargets
    ? `${fillTotalTracks} дор. на ${fillTargets.length} ` +
      `${fillTargets.length === 1 ? "день" : "днях"}: ` +
      fillTargets.map((t) => `${formatRu(t.date)} (${t.tracks})`).join(", ") +
      ". Лишние плиты остаются «в производстве»."
    : undefined;

  const handleSubmit = (order: "asc" | "desc" = "asc") => {
    const selectedKpIds = Object.entries(selectedPlatesByKp)
      .filter(([, ids]) => ids.length > 0)
      .map(([kpId]) => Number(kpId));

    let partialPlateIds: Record<number, number[]> | undefined;
    if (filterMethod === "kp" && candidatesQuery.data) {
      const candidatesByKp = new Map(
        candidatesQuery.data.items.map((kp) => [kp.kp_id, kp]),
      );
      const partial: Record<number, number[]> = {};
      for (const [kpIdStr, plateIds] of Object.entries(selectedPlatesByKp)) {
        const kpId = Number(kpIdStr);
        const kp = candidatesByKp.get(kpId);
        if (!kp || plateIds.length === 0) {
          continue;
        }
        if (plateIds.length < kp.plates.length) {
          partial[kpId] = plateIds;
        }
      }
      if (Object.keys(partial).length > 0) {
        partialPlateIds = partial;
      }
    }

    let selectedPlateQty: Record<number, Record<number, number>> | undefined;
    if (filterMethod === "kp" && candidatesQuery.data) {
      const candidatesByKp = new Map(
        candidatesQuery.data.items.map((kp) => [kp.kp_id, kp]),
      );
      const qtyPayload: Record<number, Record<number, number>> = {};
      for (const [kpIdStr, plateIds] of Object.entries(selectedPlatesByKp)) {
        if (plateIds.length === 0) {
          continue;
        }
        const kpId = Number(kpIdStr);
        const kp = candidatesByKp.get(kpId);
        if (!kp) {
          continue;
        }
        const perPlate: Record<number, number> = {};
        for (const plateId of plateIds) {
          const plate = kp.plates.find((p) => p.id === plateId);
          if (!plate) {
            continue;
          }
          perPlate[plateId] =
            selectedPlateQtyByKp[kpId]?.[plateId] ?? plate.qty;
        }
        if (Object.keys(perPlate).length > 0) {
          qtyPayload[kpId] = perPlate;
        }
      }
      if (Object.keys(qtyPayload).length > 0) {
        selectedPlateQty = qtyPayload;
      }
    }

    const fillStart = fillTargets ? fillTargets[0].date : startDate;
    const fillTracks = fillTargets
      ? Math.max(...fillTargets.map((t) => t.tracks))
      : tracksCount;

    const activePlanId = plansQuery.data?.active_plan_id ?? undefined;
    const activePlanVersion =
      activePlanId &&
      plansQuery.data?.plans.find((plan) => plan.id === activePlanId)?.version;

    buildMutation.mutate(
      {
        start_date: fillStart,
        tracks_count: fillTracks,
        filter_method: filterMethod,
        selected_kp_ids: filterMethod === "kp" ? selectedKpIds : undefined,
        selected_plate_ids: partialPlateIds,
        selected_plate_qty: selectedPlateQty,
        active_plan_id: activePlanId,
        expected_version:
          typeof activePlanVersion === "number" ? activePlanVersion : undefined,
        plan_name: planName.trim() ? planName.trim() : undefined,
        fill_targets: fillTargets ?? undefined,
        layout_reinforcement_order: order,
      },
      {
        onSuccess: () => {
          setStep(1);
          setSelectedPlatesByKp({});
          setSelectedPlateQtyByKp({});
          setExpandedKpIds(new Set());
          setPlanName("");
          setFillTargets(null);
          onCreated?.();
        },
      },
    );
  };

  const handleCancelFill = () => {
    setFillTargets(null);
    setStep(1);
    setSelectedPlatesByKp({});
    setSelectedPlateQtyByKp({});
    setExpandedKpIds(new Set());
    onCancelFill?.();
  };

  const cardTitle = isFillMode ? "Дозаполнение дней" : "Начать планирование";
  const cardSubtitle = isFillMode
    ? fillSubtitle
    : "Мастер создания нового производственного плана в три шага.";

  const buildErrorMessage =
    buildMutation.isError && !isPlanVersionConflict(buildMutation.error)
      ? getErrorMessage(buildMutation.error)
      : null;

  return {
    step,
    setStep,
    startDate,
    setStartDate,
    tracksCount,
    setTracksCount,
    filterMethod,
    setFilterMethod,
    planName,
    setPlanName,
    calendarMonth,
    setCalendarMonth,
    selectedPlatesByKp,
    selectedPlateQtyByKp,
    expandedKpIds,
    isFillMode,
    daysInfo,
    holidays,
    extraWorkdays,
    occupiedOnStart,
    maxPerDay,
    freeOnStart,
    calendarLoading: calendarQuery.isLoading || workCalendar.isLoading,
    canProceedStep1,
    canProceedStep2,
    canSubmit,
    selectionEstimate,
    estimateByKpId,
    tracksPerDay,
    tracksPerDaySource,
    candidatesQuery,
    buildMutation,
    buildErrorMessage,
    cardTitle,
    cardSubtitle,
    toggleKp,
    togglePlate,
    setPlateQty,
    toggleExpand,
    handleSubmit,
    handleCancelFill,
  };
};
