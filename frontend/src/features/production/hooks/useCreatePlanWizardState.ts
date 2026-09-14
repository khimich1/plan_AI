import { useEffect, useMemo, useState } from "react";
import { getErrorMessage } from "@/shared/lib/apiError";
import { isPlanVersionConflict } from "@/shared/lib/planConflict";
import {
  useAnalyzeSubstratesMutation,
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
  AnalyzeSubstratesResponse,
  CapacityOption,
  FillTargetItem,
  FilterMethod,
  KpCandidateItem,
  KpCandidatePlateItem,
  PendingPromiseExclusion,
  PromisedBlockItem,
  PromiseExclusion,
  SgpReservationItem,
  SubstrateRecommendation,
  UrgentPosition,
} from "@/features/production/types/production";
import { formatRu } from "@/features/production/components/create-plan-wizard/utils";

export const isoWeekStart = (isoDate: string): string => {
  const [y, m, d] = isoDate.split("-").map(Number);
  const date = new Date(y, (m || 1) - 1, d || 1);
  const day = date.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + mondayOffset);
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
};

export type UseCreatePlanWizardStateOptions = {
  onCreated?: () => void;
  fillRequest?: FillTargetItem[] | null;
  onFillRequestConsumed?: () => void;
  onCancelFill?: () => void;
};

const fillTargetsKey = (targets: FillTargetItem[] | null): string =>
  targets?.map((t) => `${t.date}:${t.tracks}`).join("|") ?? "";

const maxFillDate = (targets: FillTargetItem[]): string =>
  targets.reduce((max, t) => (t.date > max ? t.date : max), targets[0].date);

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
  const [analyzeResult, setAnalyzeResult] =
    useState<AnalyzeSubstratesResponse | null>(null);
  const [exclusionByKp, setExclusionByKp] = useState<
    Record<number, PromiseExclusion>
  >({});
  const [pendingExclusion, setPendingExclusion] =
    useState<PendingPromiseExclusion | null>(null);

  const calendarQuery = useGlobalCalendarQuery();
  const plansQuery = usePlansListQuery();
  const candidatesQuery = useKpCandidatesQuery(true);
  const freePlatesQuery = useSgpFreePlatesQuery(true);
  const buildMutation = useBuildPlanMutation();
  const analyzeMutation = useAnalyzeSubstratesMutation();

  const daysInfo = calendarQuery.data?.days_info ?? {};
  const targetsKey = fillTargetsKey(fillTargets);

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

  /** Выбор плиты по id+kp+qty без полного KpCandidateItem (срочные / подложки). */
  const setPlateSelectionById = (
    kpId: number,
    plateId: number,
    qty: number,
    selected: boolean,
  ) => {
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      const current = next[kpId] ?? [];
      if (selected) {
        if (!current.includes(plateId)) {
          next[kpId] = [...current, plateId];
        }
      } else {
        const filtered = current.filter((id) => id !== plateId);
        if (filtered.length === 0) {
          delete next[kpId];
        } else {
          next[kpId] = filtered;
        }
      }
      return next;
    });
    setSelectedPlateQtyByKp((prev) => {
      const next = { ...prev };
      if (selected) {
        next[kpId] = { ...(next[kpId] ?? {}), [plateId]: qty };
      } else {
        const per = { ...(next[kpId] ?? {}) };
        delete per[plateId];
        if (Object.keys(per).length === 0) {
          delete next[kpId];
        } else {
          next[kpId] = per;
        }
      }
      return next;
    });
  };

  const applyUrgentDefaults = (positions: UrgentPosition[]) => {
    if (positions.length === 0) return;
    setFilterMethod("kp");
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      for (const pos of positions) {
        const current = next[pos.kp_id] ?? [];
        if (!current.includes(pos.plate_id)) {
          next[pos.kp_id] = [...current, pos.plate_id];
        }
      }
      return next;
    });
    setSelectedPlateQtyByKp((prev) => {
      const next = { ...prev };
      for (const pos of positions) {
        next[pos.kp_id] = {
          ...(next[pos.kp_id] ?? {}),
          [pos.plate_id]: pos.qty_remaining,
        };
      }
      return next;
    });
  };

  const applySubstrateDefaults = (recommendations: SubstrateRecommendation[]) => {
    if (recommendations.length === 0) return;
    setFilterMethod("kp");
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      for (const rec of recommendations) {
        const current = next[rec.kp_id] ?? [];
        if (!current.includes(rec.plate_id)) {
          next[rec.kp_id] = [...current, rec.plate_id];
        }
      }
      return next;
    });
    setSelectedPlateQtyByKp((prev) => {
      const next = { ...prev };
      for (const rec of recommendations) {
        next[rec.kp_id] = {
          ...(next[rec.kp_id] ?? {}),
          [rec.plate_id]: rec.qty_recommended,
        };
      }
      return next;
    });
  };

  const {
    mutate: analyzeSubstrates,
    reset: resetAnalyze,
    isPending: analyzePending,
    isError: analyzeIsError,
    error: analyzeError,
  } = analyzeMutation;

  /** Общий analyze: авто при fillTargets и кнопка «Найти подложки». */
  const runAnalyzeSubstrates = (opts?: { isCancelled?: () => boolean }) => {
    if (!fillTargets || fillTargets.length === 0) return;
    const deadline_until = maxFillDate(fillTargets);
    analyzeSubstrates(
      { fill_targets: fillTargets, deadline_until },
      {
        onSuccess: (data) => {
          if (opts?.isCancelled?.()) return;
          setAnalyzeResult(data);
          applyUrgentDefaults(data.urgent_positions);
          applySubstrateDefaults(data.substrate_recommendations);
        },
      },
    );
  };

  /** Автопредложение срочных/подложек: один analyze на набор fillTargets. */
  useEffect(() => {
    if (!fillTargets || fillTargets.length === 0) {
      setAnalyzeResult(null);
      return;
    }

    let cancelled = false;
    setAnalyzeResult(null);
    runAnalyzeSubstrates({ isCancelled: () => cancelled });

    return () => {
      cancelled = true;
    };
    // runAnalyzeSubstrates / fillTargets captured for current targetsKey
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [targetsKey, analyzeSubstrates]);

  const toggleUrgentPosition = (pos: UrgentPosition) => {
    const selected = (selectedPlatesByKp[pos.kp_id] ?? []).includes(pos.plate_id);
    setFilterMethod("kp");
    setPlateSelectionById(pos.kp_id, pos.plate_id, pos.qty_remaining, !selected);
  };

  /**
   * Применяет выбранный option из capacity_deficit к fill_targets.
   * Не вызывает PUT /day-capacity — только bump fill или добавление дня.
   */
  const applyCapacityOption = (option: CapacityOption) => {
    if (!fillTargets || option.add_tracks <= 0) return;
    const { date, add_tracks: addTracks } = option;

    setFillTargets((prev) => {
      if (!prev) return prev;
      const hasDate = prev.some((t) => t.date === date);
      if (!hasDate) {
        return [...prev, { date, tracks: addTracks }].sort((a, b) =>
          a.date.localeCompare(b.date),
        );
      }
      return prev.map((t) =>
        t.date === date ? { ...t, tracks: t.tracks + addTracks } : t,
      );
    });
  };

  const toggleSubstrateRecommendation = (rec: SubstrateRecommendation) => {
    const selected = (selectedPlatesByKp[rec.kp_id] ?? []).includes(rec.plate_id);
    setFilterMethod("kp");
    setPlateSelectionById(
      rec.kp_id,
      rec.plate_id,
      rec.qty_recommended,
      !selected,
    );
  };

  const defaultQtyMap = (kp: KpCandidateItem, plateIds: number[]) => {
    const map: Record<number, number> = {};
    for (const plate of kp.plates) {
      if (plateIds.includes(plate.id)) {
        map[plate.id] = plate.qty;
      }
    }
    return map;
  };

  const applyPromiseDefaults = (items: KpCandidateItem[]) => {
    const selectedWeeks = new Set(
      (fillTargets ?? []).map((target) => isoWeekStart(target.date)),
    );
    const promised = items.filter((kp) => {
      if (!kp.promise || kp.plates.length === 0) {
        return false;
      }
      return (
        kp.promise.status === "overdue" ||
        selectedWeeks.has(kp.promise.week_start)
      );
    });
    if (promised.length === 0) {
      return;
    }
    setFilterMethod("kp");
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      let changed = false;
      for (const kp of promised) {
        if (exclusionByKp[kp.kp_id]) {
          continue;
        }
        const ids = kp.plates.map((plate) => plate.id);
        const already = next[kp.kp_id] ?? [];
        const merged = [...new Set([...already, ...ids])];
        if (
          already.length !== merged.length ||
          merged.some((id, index) => id !== already[index])
        ) {
          next[kp.kp_id] = merged;
          changed = true;
        }
      }
      return changed ? next : prev;
    });
    setSelectedPlateQtyByKp((prev) => {
      const next = { ...prev };
      let changed = false;
      for (const kp of promised) {
        if (exclusionByKp[kp.kp_id]) {
          continue;
        }
        const ids = kp.plates.map((plate) => plate.id);
        const mapped = defaultQtyMap(kp, ids);
        next[kp.kp_id] = { ...(next[kp.kp_id] ?? {}), ...mapped };
        changed = true;
      }
      return changed ? next : prev;
    });
  };

  useEffect(() => {
    const items = candidatesQuery.data?.items;
    if (!items || !fillTargets || fillTargets.length === 0) {
      return;
    }
    applyPromiseDefaults(items);
    // applyPromiseDefaults reads fillTargets / candidates for the current targetsKey
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [candidatesQuery.data, targetsKey, exclusionByKp]);

  const promisedBlockItems = useMemo((): PromisedBlockItem[] => {
    const payload = candidatesQuery.data;
    if (!payload) {
      return [];
    }
    const selectedWeeks = new Set(
      (fillTargets ?? []).map((target) => isoWeekStart(target.date)),
    );
    const byId = new Map(payload.items.map((kp) => [kp.kp_id, kp]));
    const seen = new Set<number>();
    const result: PromisedBlockItem[] = [];

    const pushItem = (
      kpId: number,
      weekStart: string,
      promisedDate: string,
      tracks: number,
      status: PromisedBlockItem["status"],
    ) => {
      if (seen.has(kpId)) {
        return;
      }
      if (status !== "overdue" && !selectedWeeks.has(weekStart)) {
        return;
      }
      seen.add(kpId);
      result.push({
        kp_id: kpId,
        promised_date: promisedDate,
        tracks,
        status,
        week_start: weekStart,
        customer_name: byId.get(kpId)?.customer_name,
      });
    };

    const weeks = [...(payload.promised_weeks ?? [])].sort((left, right) => {
      const leftOverdue = left.items.some((item) => item.status === "overdue")
        ? 0
        : 1;
      const rightOverdue = right.items.some((item) => item.status === "overdue")
        ? 0
        : 1;
      if (leftOverdue !== rightOverdue) {
        return leftOverdue - rightOverdue;
      }
      return left.week_start.localeCompare(right.week_start);
    });
    for (const week of weeks) {
      for (const item of week.items) {
        pushItem(
          item.kp_id,
          week.week_start,
          item.promised_date,
          item.tracks,
          item.status,
        );
      }
    }
    for (const kp of payload.items) {
      if (!kp.promise) {
        continue;
      }
      pushItem(
        kp.kp_id,
        kp.promise.week_start,
        kp.promise.promised_date,
        kp.promise.tracks,
        kp.promise.status,
      );
    }
    return result;
  }, [candidatesQuery.data, fillTargets]);

  const exclusions = useMemo(
    () => Object.values(exclusionByKp),
    [exclusionByKp],
  );

  const deselectKp = (kpId: number) => {
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      delete next[kpId];
      return next;
    });
    setSelectedPlateQtyByKp((qty) => {
      const next = { ...qty };
      delete next[kpId];
      return next;
    });
  };

  const deselectPlate = (kpId: number, plateId: number) => {
    setSelectedPlatesByKp((prev) => {
      const next = { ...prev };
      const current = next[kpId] ?? [];
      const filtered = current.filter((id) => id !== plateId);
      if (filtered.length === 0) {
        delete next[kpId];
      } else {
        next[kpId] = filtered;
      }
      return next;
    });
    setSelectedPlateQtyByKp((qty) => {
      const perKp = { ...(qty[kpId] ?? {}) };
      delete perKp[plateId];
      if (Object.keys(perKp).length === 0) {
        const next = { ...qty };
        delete next[kpId];
        return next;
      }
      return { ...qty, [kpId]: perKp };
    });
  };

  const toggleKp = (kp: KpCandidateItem) => {
    const selected = (selectedPlatesByKp[kp.kp_id] ?? []).length > 0;
    if (selected && kp.promise && !exclusionByKp[kp.kp_id]) {
      setPendingExclusion({
        kpId: kp.kp_id,
        weekStart: kp.promise.week_start,
        kind: "whole",
      });
      return;
    }
    if (!selected && exclusionByKp[kp.kp_id]) {
      setExclusionByKp((prev) => {
        const next = { ...prev };
        delete next[kp.kp_id];
        return next;
      });
    }
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
    const currentIds = selectedPlatesByKp[kp.kp_id];
    const isRemoving = currentIds !== undefined && currentIds.includes(plateId);
    if (isRemoving && kp.promise && !exclusionByKp[kp.kp_id]) {
      setPendingExclusion({
        kpId: kp.kp_id,
        weekStart: kp.promise.week_start,
        kind: "partial",
        plateId,
      });
      return;
    }
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
    !buildMutation.isSuccess &&
    pendingExclusion === null;

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
    setAnalyzeResult(null);
    setFilterMethod("all");
    setExclusionByKp({});
    setPendingExclusion(null);
    resetAnalyze();
  };

  const confirmExclusion = (reason: string) => {
    const trimmed = reason.trim();
    if (!pendingExclusion || !trimmed) {
      return;
    }
    const pending = pendingExclusion;
    setExclusionByKp((prev) => ({
      ...prev,
      [pending.kpId]: {
        kp_id: pending.kpId,
        week_start: pending.weekStart,
        reason: trimmed,
      },
    }));
    setPendingExclusion(null);
    if (pending.kind === "whole") {
      deselectKp(pending.kpId);
      return;
    }
    if (pending.plateId != null) {
      deselectPlate(pending.kpId, pending.plateId);
    }
  };

  const cancelExclusion = () => setPendingExclusion(null);

  const togglePromisedKp = (kpId: number) => {
    const kp = candidatesQuery.data?.items.find((item) => item.kp_id === kpId);
    if (kp) {
      toggleKp(kp);
    }
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
    if (pendingExclusion) return;

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
        exclusions:
          filterMethod === "kp" && exclusions.length > 0
            ? exclusions
            : undefined,
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

  const analyzeErrorMessage = analyzeIsError
    ? getErrorMessage(analyzeError)
    : null;

  /** HTTP-ошибка analyze или error_message из analysis_meta (status=error). */
  const substrateErrorMessage =
    analyzeErrorMessage ??
    (analyzeResult?.analysis_meta.optimization_status === "error"
      ? analyzeResult.analysis_meta.error_message ||
        "Ошибка анализа подложек"
      : null);

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
    fillTargets,
    analyzeResult,
    analyzePending,
    analyzeErrorMessage,
    substrateErrorMessage,
    runAnalyzeSubstrates,
    toggleUrgentPosition,
    toggleSubstrateRecommendation,
    setPlateSelectionById,
    applyCapacityOption,
    promisedBlockItems,
    pendingExclusion,
    exclusions,
    confirmExclusion,
    cancelExclusion,
    togglePromisedKp,
  };
};
