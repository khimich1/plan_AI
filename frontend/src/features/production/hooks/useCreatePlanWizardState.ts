import { useEffect, useMemo, useState } from "react";
import { getErrorMessage } from "@/shared/lib/apiError";
import { isPlanVersionConflict } from "@/shared/lib/planConflict";
import {
  useBuildPlanMutation,
  useGlobalCalendarQuery,
  useKpCandidatesQuery,
  usePlansListQuery,
} from "@/features/production/hooks/useProductionQueries";
import { useSgpFreePlatesQuery } from "@/features/production/hooks/useSgpQueries";
import {
  freeQtyForPlate,
  pickFreeReservations,
} from "@/features/production/lib/sgpFreeMatch";
import {
  getBasketKind,
  type BasketDayKind,
} from "@/features/production/lib/basketDayKind";
import { planNameFromDates } from "@/features/production/lib/planNameFromDates";
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
  KpCandidatePlateItem,
  SgpReservationItem,
} from "@/features/production/types/production";
import { formatRu } from "@/features/production/components/create-plan-wizard/utils";

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
  const [filterMethod, setFilterMethod] = useState<FilterMethod>("all");
  const [selectedPlatesByKp, setSelectedPlatesByKp] = useState<
    Record<number, number[]>
  >({});
  const [selectedPlateQtyByKp, setSelectedPlateQtyByKp] = useState<
    Record<number, Record<number, number>>
  >({});
  const [expandedKpIds, setExpandedKpIds] = useState<Set<number>>(new Set());
  const [planName, setPlanName] = useState<string>("");
  const [fillTargets, setFillTargets] = useState<FillTargetItem[] | null>(null);
  const [basketKind, setBasketKind] = useState<BasketDayKind | null>(null);
  const [sgpReservations, setSgpReservations] = useState<SgpReservationItem[]>([]);
  const [pendingClose, setPendingClose] = useState<{
    kp: KpCandidateItem;
    plate: KpCandidatePlateItem;
    freeQty: number;
    closeQty: number;
  } | null>(null);

  const calendarQuery = useGlobalCalendarQuery();
  const plansQuery = usePlansListQuery();
  const candidatesQuery = useKpCandidatesQuery(true);
  const freePlatesQuery = useSgpFreePlatesQuery(true);
  const buildMutation = useBuildPlanMutation();

  const daysInfo = calendarQuery.data?.days_info ?? {};

  useEffect(() => {
    if (fillRequest && fillRequest.length > 0) {
      setFillTargets(fillRequest);
      setPlanName(planNameFromDates(fillRequest.map((t) => t.date)));
      onFillRequestConsumed?.();
    }
  }, [fillRequest, onFillRequestConsumed]);

  useEffect(() => {
    if (fillTargets && fillTargets.length > 0) {
      setBasketKind(getBasketKind(fillTargets, daysInfo));
    }
  }, [fillTargets, daysInfo]);

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

  const isFillMode = fillTargets !== null && fillTargets.length > 0;

  const tracksPerDay =
    isFillMode && fillTargets
      ? Math.max(...fillTargets.map((t) => t.tracks))
      : 1;
  const tracksPerDaySource = "календарь" as const;

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

  const hasAnyPlateSelected =
    filterMethod === "all" ||
    Object.values(selectedPlatesByKp).some((ids) => ids.length > 0) ||
    sgpReservations.length > 0;
  const canSubmit =
    isFillMode &&
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

  const freeItems = freePlatesQuery.data?.items ?? [];

  const freeQtyByPlateKey = useMemo(() => {
    const map = new Map<string, number>();
    const candidates = candidatesQuery.data?.items ?? [];
    for (const kp of candidates) {
      for (const plate of kp.plates) {
        const key = `${kp.kp_id}:${plate.id}`;
        map.set(key, freeQtyForPlate(plate, freeItems));
      }
    }
    return map;
  }, [candidatesQuery.data, freeItems]);

  const resetSelection = () => {
    setSelectedPlatesByKp({});
    setSelectedPlateQtyByKp({});
    setExpandedKpIds(new Set());
    setPlanName("");
    setFillTargets(null);
    setBasketKind(null);
    setSgpReservations([]);
    setPendingClose(null);
  };

  const proposeCloseFromSgp = (kp: KpCandidateItem, plate: KpCandidatePlateItem) => {
    const freeQty = freeQtyByPlateKey.get(`${kp.kp_id}:${plate.id}`) ?? 0;
    const selectedQty =
      selectedPlateQtyByKp[kp.kp_id]?.[plate.id] ?? plate.qty;
    const closeQty = Math.min(freeQty, selectedQty, plate.qty);
    if (closeQty <= 0) return;
    setPendingClose({ kp, plate, freeQty, closeQty });
  };

  const confirmCloseFromSgp = () => {
    if (!pendingClose) return;
    const { kp, plate, closeQty } = pendingClose;
    const picks = pickFreeReservations(plate, freeItems, closeQty);
    if (picks.length === 0) {
      setPendingClose(null);
      return;
    }
    const reservedQty = picks.reduce((s, p) => s + p.qty, 0);
    setSgpReservations((prev) => [
      ...prev,
      ...picks.map((p) => ({
        sgp_id: p.sgp_id,
        target_kp_id: kp.kp_id,
        qty: p.qty,
      })),
    ]);
    // If manual KP selection is active, reduce selected production qty.
    if (filterMethod === "kp") {
      setSelectedPlateQtyByKp((q) => {
        const current = q[kp.kp_id]?.[plate.id] ?? plate.qty;
        const nextQty = Math.max(0, current - reservedQty);
        const kpMap = { ...(q[kp.kp_id] ?? {}) };
        if (nextQty > 0) {
          kpMap[plate.id] = nextQty;
        } else {
          delete kpMap[plate.id];
        }
        return { ...q, [kp.kp_id]: kpMap };
      });
      setSelectedPlatesByKp((prev) => {
        const current = prev[kp.kp_id] ?? [];
        const nextQty =
          (selectedPlateQtyByKp[kp.kp_id]?.[plate.id] ?? plate.qty) - reservedQty;
        let nextIds = current.includes(plate.id)
          ? [...current]
          : [...current, plate.id];
        if (nextQty <= 0) {
          nextIds = nextIds.filter((id) => id !== plate.id);
        }
        return { ...prev, [kp.kp_id]: nextIds };
      });
    }
    setPendingClose(null);
  };

  const cancelCloseFromSgp = () => setPendingClose(null);

  const handleSubmit = (order: "asc" | "desc" = "asc") => {
    if (!fillTargets || fillTargets.length === 0) return;

    const selectedKpIds = Array.from(
      new Set([
        ...Object.entries(selectedPlatesByKp)
          .filter(([, ids]) => ids.length > 0)
          .map(([kpId]) => Number(kpId)),
        ...sgpReservations.map((r) => r.target_kp_id),
      ]),
    );

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

    const autoName = planNameFromDates(fillTargets.map((t) => t.date));
    const fillStart = fillTargets[0].date;
    const fillTracks = Math.max(...fillTargets.map((t) => t.tracks));

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
        plan_name: autoName,
        fill_targets: fillTargets,
        layout_reinforcement_order: order,
        sgp_reservations:
          sgpReservations.length > 0 ? sgpReservations : undefined,
      },
      {
        onSuccess: () => {
          resetSelection();
          onCreated?.();
        },
      },
    );
  };

  const handleCancelFill = () => {
    resetSelection();
    onCancelFill?.();
  };

  const cardTitle =
    basketKind === "empty" ? "Начать планирование" : "Дозаполнение дней";
  const cardSubtitle = fillSubtitle;

  const buildErrorMessage =
    buildMutation.isError && !isPlanVersionConflict(buildMutation.error)
      ? getErrorMessage(buildMutation.error)
      : null;

  return {
    filterMethod,
    setFilterMethod,
    planName,
    selectedPlatesByKp,
    selectedPlateQtyByKp,
    expandedKpIds,
    isFillMode,
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
    basketKind,
    freeQtyByPlateKey,
    sgpReservations,
    pendingClose,
    proposeCloseFromSgp,
    confirmCloseFromSgp,
    cancelCloseFromSgp,
    toggleKp,
    togglePlate,
    setPlateQty,
    toggleExpand,
    handleSubmit,
    handleCancelFill,
  };
};
