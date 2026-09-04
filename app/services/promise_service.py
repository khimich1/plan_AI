"""Promise quote and holds. Occupancy is fail-closed; holds do not reduce free."""

from __future__ import annotations

import logging
from collections.abc import Callable, Collection, Mapping, Sequence
from dataclasses import dataclass
from datetime import date, datetime, time
from typing import Any

from app.core.settings import get_settings
from app.repositories.kp_archive_repository import KpArchiveRepository
from app.repositories.plan_repository import PlanRepository
from app.repositories.promise_repository import PromiseRepository
from app.schemas.archive import (
    PromiseHoldAllocation,
    PromiseHoldResponse,
    PromiseQuoteResponse,
    PromiseQuoteWeek,
    PromiseQuoteWindow,
    PromiseTracksPerDayResponse,
)
from app.security.offer_access import (
    assert_offer_read_access,
    assert_offer_write_access,
    is_admin,
)
from app.services.plan_distribution_service import PlanDistributionService
from core.kp.offers_write import MovePromisePayload, commit_move_to_production
from core.production.capacity import TRACKS_PER_DAY_HARD_CAP
from core.production.promise_buckets import (
    OccupancyUnavailableError,
    PromiseQuote,
    build_quote,
    build_weeks,
    workday_predicate,
)
from core.work_calendar import load_extra_workdays, load_holidays

logger = logging.getLogger(__name__)

_OCCUPANCY_UNAVAILABLE = (
    "Недоступна занятость плана — котировка остановлена (fail-closed)."
)

OccupancyLoader = Callable[[], Mapping[date | str, int] | None]
CalendarLoader = Callable[[], object]
KpLoader = Callable[[int], dict | None]
WorkdayFn = Callable[[date], bool]


class PromiseError(Exception):
    """Base error for promise quote and holds."""


class PromiseNotFoundError(PromiseError):
    """КП не найдено."""


class PromiseHoldNotFoundError(PromiseError):
    """Активный холд не найден."""


class PromiseHoldForbiddenError(PromiseError):
    """Снять холд может только владелец или администратор."""


class PromiseHoldUnavailableError(PromiseError):
    """Нет окна корзин — холд закрепить нельзя."""


class PromiseGateError(PromiseError):
    """Requested move date is earlier than the recalculated promised_date."""

    def __init__(self, message: str, *, earliest: date | None = None) -> None:
        super().__init__(message)
        self.earliest = earliest


class PromiseExclusionError(PromiseError):
    """Exclusion is missing a required reason (level 2 still allows the build)."""


class PromiseKnobInvalidError(PromiseError):
    """tracks_per_day must be 1..TRACKS_PER_DAY_HARD_CAP."""


@dataclass(frozen=True, slots=True)
class PlanExclusionRecord:
    """One journal row written with the plan commit (optional notification)."""

    exclusion_id: int
    notification_id: int | None
    kp_id: int
    week_start: date
    reason: str


@dataclass(frozen=True, slots=True)
class PlanCommitSettlement:
    """Result of consuming / marking overdue allocations after a plan commit."""

    consumed_alloc_ids: tuple[int, ...]
    overdue_alloc_ids: tuple[int, ...]
    consumed_promise_ids: tuple[int, ...]


KIND_PROMISED_DATE_SHIFTED = "promised_date_shifted"


@dataclass(frozen=True, slots=True)
class PromiseRecalcResult:
    """Outcome of rewriting an active promise after KP composition edit."""

    promise_id: int | None
    old_promised_date: date | None
    new_promised_date: date | None
    tracks: int
    notified: bool
    notification_id: int | None


def _occupied_int(info: object) -> int:
    raw = info.get("occupied", 0) if isinstance(info, Mapping) else info
    try:
        return int(raw)
    except (TypeError, ValueError) as exc:
        raise OccupancyUnavailableError(_OCCUPANCY_UNAVAILABLE) from exc


def occupancy_map_from_calendar(calendar: object) -> dict[str, int]:
    """Convert PlanDistributionService calendar to date→occupied.

    ``None`` calendar is a successful empty factory (no plans), not a loader
    failure. Malformed ``days_info`` is fail-closed.
    """
    if calendar is None:
        return {}
    if not isinstance(calendar, Mapping):
        raise OccupancyUnavailableError(_OCCUPANCY_UNAVAILABLE)
    days_info = calendar.get("days_info")
    if days_info is None:
        return {}
    if not isinstance(days_info, Mapping):
        raise OccupancyUnavailableError(_OCCUPANCY_UNAVAILABLE)
    return {str(day): _occupied_int(info) for day, info in days_info.items()}


def _default_calendar(db_path: str | None = None) -> object:
    return PlanDistributionService().get_global_calendar_info(
        PlanRepository(db_path=db_path)
    )


def load_plan_occupancy(
    *,
    calendar_loader: CalendarLoader | None = None,
    db_path: str | None = None,
) -> dict[str, int]:
    """Load plan occupancy. Loader exceptions become OccupancyUnavailableError."""
    loader = calendar_loader or (lambda: _default_calendar(db_path))
    try:
        calendar = loader()
    except OccupancyUnavailableError:
        raise
    except Exception as exc:
        logger.exception("promise quote: occupancy loader failed")
        raise OccupancyUnavailableError(_OCCUPANCY_UNAVAILABLE) from exc
    return occupancy_map_from_calendar(calendar)


def _total_length_m(raw: dict) -> float:
    plates = raw.get("plates") or []
    total = 0.0
    for plate in plates:
        try:
            length_m = float(plate.get("length_m") or 0)
            qty = float(plate.get("qty") or 0)
        except (TypeError, ValueError):
            continue
        if length_m <= 0 or qty <= 0:
            continue
        total += length_m * qty
    return total


def _quote_to_response(quote: PromiseQuote) -> PromiseQuoteResponse:
    window = None
    if quote.window is not None:
        window = PromiseQuoteWindow(
            from_week=quote.window.from_week,
            to_week=quote.window.to_week,
            promised_date=quote.window.promised_date,
        )
    return PromiseQuoteResponse(
        tracks=quote.tracks,
        solo_days=quote.solo_days,
        solo_date=quote.solo_date,
        solo_week_end_date=quote.solo_week_end_date,
        earliest_start_week=quote.earliest_start_week,
        window=window,
        weeks=[
            PromiseQuoteWeek(
                week_start=week.week_start,
                workdays=week.workdays,
                capacity=week.capacity,
                planned=week.planned,
                promised=week.promised,
                held=week.held,
                free=week.free,
            )
            for week in quote.weeks
        ],
        knob=quote.knob,
    )


def _actor(user: dict) -> str:
    username = user.get("username")
    if username:
        return str(username)
    user_id = user.get("id")
    return "" if user_id is None else str(user_id)


def _can_release_hold(user: dict, hold: Mapping) -> bool:
    if is_admin(user):
        return True
    actor = _actor(user)
    return bool(actor) and str(hold.get("created_by") or "") == actor


def _as_date(value: date | datetime | str) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])


def _as_dt(value: datetime | str) -> datetime:
    if isinstance(value, datetime):
        return value
    return datetime.fromisoformat(str(value))


def _iso_dt(value: datetime) -> str:
    return value.isoformat(timespec="seconds")


def _parse_optional_dt(value: object) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    if not text:
        return None
    return datetime.fromisoformat(text.replace(" ", "T", 1))


def _knob_to_response(row: Mapping) -> PromiseTracksPerDayResponse:
    return PromiseTracksPerDayResponse(
        tracks_per_day=int(row["tracks_per_day"]),
        updated_by=row.get("updated_by"),
        updated_at=_parse_optional_dt(row.get("updated_at")),
        min=1,
        max=TRACKS_PER_DAY_HARD_CAP,
    )


def _hold_to_response(row: Mapping) -> PromiseHoldResponse:
    expires_raw = row.get("expires_at")
    if expires_raw is None:
        raise PromiseError("hold is missing expires_at")
    return PromiseHoldResponse(
        id=int(row["id"]),
        kp_id=int(row["kp_id"]),
        kind="hold",
        status=row["status"],
        tracks_total=int(row["tracks_total"]),
        promised_date=_as_date(row["promised_date"]),
        expires_at=_as_dt(expires_raw),
        created_by=row.get("created_by"),
        created_at=_as_dt(row["created_at"]),
        allocations=[
            PromiseHoldAllocation(week_start=_as_date(week_start), tracks=int(tracks))
            for week_start, tracks in row.get("allocations") or ()
        ],
    )


def end_of_local_day(day: date) -> datetime:
    return datetime.combine(day, time(23, 59, 59))


def _normalized_exclusions(
    exclusions: Sequence[Mapping[str, Any] | Any],
) -> list[tuple[int, date, str]]:
    unique: dict[tuple[int, date], str] = {}
    for item in exclusions:
        if isinstance(item, Mapping):
            kp_id = item.get("kp_id")
            week_start = item.get("week_start")
            reason = item.get("reason")
        else:
            kp_id = getattr(item, "kp_id", None)
            week_start = getattr(item, "week_start", None)
            reason = getattr(item, "reason", None)
        text = "" if reason is None else str(reason).strip()
        if not text:
            raise PromiseExclusionError("Причина снятия обещанного КП обязательна")
        if kp_id is None or week_start is None:
            raise PromiseExclusionError("Исключение должно содержать kp_id и week_start")
        unique[(int(kp_id), _as_date(week_start))] = text
    return [(kp_id, week, reason) for (kp_id, week), reason in unique.items()]


class PromiseService:
    """Quote weekly buckets and pin a same-day hold (not a KP status)."""

    def __init__(
        self,
        *,
        db_path: str | None = None,
        repository: PromiseRepository | None = None,
        kp_loader: KpLoader | None = None,
        occupancy_loader: OccupancyLoader | None = None,
        today: date | None = None,
        now: datetime | None = None,
        is_workday: WorkdayFn | None = None,
        week_count: int = 12,
    ) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)
        self.repository = repository or PromiseRepository(db_path=self.db_path)
        self._kp_loader = kp_loader
        self._occupancy_loader = occupancy_loader
        self._now = now
        self._today = today
        self._is_workday = is_workday
        self._week_count = week_count

    def get_quote(self, kp_id: int, *, user: dict) -> PromiseQuoteResponse:
        raw = self._load_kp(kp_id)
        assert_offer_read_access(user, raw)
        return _quote_to_response(self._compute_quote(raw, exclude_kp_id=kp_id))

    def get_tracks_per_day(self, *, user: dict) -> PromiseTracksPerDayResponse:
        del user
        return _knob_to_response(self.repository.get_promise_tracks_per_day_row())

    def set_tracks_per_day(
        self, tracks_per_day: int, *, user: dict
    ) -> PromiseTracksPerDayResponse:
        value = int(tracks_per_day)
        if value < 1 or value > TRACKS_PER_DAY_HARD_CAP:
            raise PromiseKnobInvalidError(
                f"Дорожек в день: от 1 до {TRACKS_PER_DAY_HARD_CAP}."
            )
        row = self.repository.set_promise_tracks_per_day(
            value,
            updated_by=_actor(user),
            updated_at=self._moment(),
        )
        return _knob_to_response(row)

    def create_hold(self, kp_id: int, *, user: dict) -> PromiseHoldResponse:
        raw = self._load_kp(kp_id)
        assert_offer_write_access(user, raw)
        quote = self._compute_quote(raw, exclude_kp_id=kp_id)
        if quote.tracks <= 0 or quote.window is None:
            raise PromiseHoldUnavailableError(
                "Нельзя закрепить срок: нет окна в недельных корзинах."
            )
        moment = self._moment()
        row = self.repository.insert_hold(
            kp_id=int(kp_id),
            tracks_total=quote.tracks,
            promised_date=quote.window.promised_date,
            allocations=quote.window.allocations,
            created_by=_actor(user),
            created_at=moment,
            expires_at=end_of_local_day(moment.date()),
            now=moment,
        )
        return _hold_to_response(row)

    def release_hold(self, kp_id: int, *, user: dict) -> PromiseHoldResponse:
        raw = self._load_kp(kp_id)
        assert_offer_write_access(user, raw)
        moment = self._moment()
        current = self.repository.get_active_hold(kp_id, now=moment)
        if current is None:
            raise PromiseHoldNotFoundError("Активный холд не найден.")
        if not _can_release_hold(user, current):
            raise PromiseHoldForbiddenError(
                "Снять холд может только владелец или администратор."
            )
        released = self.repository.release_hold(kp_id, now=moment)
        if released is None:
            raise PromiseHoldNotFoundError("Активный холд не найден.")
        return _hold_to_response(released)

    def get_hold(self, kp_id: int, *, user: dict) -> PromiseHoldResponse | None:
        raw = self._load_kp(kp_id)
        assert_offer_read_access(user, raw)
        row = self.repository.get_latest_hold(kp_id, now=self._moment())
        return None if row is None else _hold_to_response(row)

    def evaluate_move_gate(
        self,
        kp_id: int,
        requested_date: date,
        *,
        user: dict,
        raw: dict | None = None,
    ) -> MovePromisePayload | None:
        """Recalculate buckets; reject if requested_date < promised_date.

        Returns a write payload (or None when the KP has no tracks). Occupancy
        errors propagate as OccupancyUnavailableError — never «all free».
        """
        offer = raw if raw is not None else self._load_kp(kp_id)
        quote = self._compute_quote(offer, exclude_kp_id=int(kp_id))
        if quote.tracks <= 0:
            return None
        if quote.window is None:
            raise PromiseGateError(
                "Нельзя перевести в производство: нет окна в недельных корзинах."
            )
        earliest = quote.window.promised_date
        if requested_date < earliest:
            raise PromiseGateError(
                f"Срок раньше ближайшей возможной даты {earliest.strftime('%d.%m.%Y')}.",
                earliest=earliest,
            )
        moment = self._moment()
        hold = self.repository.get_active_hold(int(kp_id), now=moment)
        if hold is not None:
            return MovePromisePayload(
                tracks_total=int(hold["tracks_total"]),
                promised_date=_as_date(hold["promised_date"]).isoformat(),
                allocations=tuple(
                    (week.isoformat(), int(tracks))
                    for week, tracks in hold.get("allocations") or ()
                ),
                created_by=_actor(user),
                created_at=_iso_dt(moment),
                convert_hold_id=int(hold["id"]),
            )
        return MovePromisePayload(
            tracks_total=quote.tracks,
            promised_date=earliest.isoformat(),
            allocations=tuple(
                (week.isoformat(), int(tracks))
                for week, tracks in quote.window.allocations
            ),
            created_by=_actor(user),
            created_at=_iso_dt(moment),
        )

    def commit_move_with_gate(
        self,
        kp_id: int,
        execution_terms: str,
        *,
        user: dict,
        raw: dict | None = None,
    ) -> int:
        """Gate + ``commit_move_to_production`` (promise written in that tx)."""
        requested = datetime.strptime(execution_terms, "%d.%m.%Y").date()
        payload = self.evaluate_move_gate(
            kp_id, requested, user=user, raw=raw
        )
        return commit_move_to_production(
            kp_id, execution_terms, self.db_path, promise=payload
        )

    def settle_plan_commit(
        self,
        *,
        entered_kp_ids: Collection[int],
        covered_weeks: Sequence[date],
        _external_conn: Any | None = None,
    ) -> PlanCommitSettlement:
        """Consume covered allocs of entered KPs; overdue missed ones (level 2).

        Does not raise on overdue — the commit continues. Write errors propagate
        so the caller can roll back the same connection.
        """
        raw = self.repository.apply_plan_commit_settlement(
            entered_kp_ids=entered_kp_ids,
            covered_weeks=covered_weeks,
            _external_conn=_external_conn,
        )
        return PlanCommitSettlement(
            consumed_alloc_ids=raw["consumed_alloc_ids"],
            overdue_alloc_ids=raw["overdue_alloc_ids"],
            consumed_promise_ids=raw["consumed_promise_ids"],
        )

    def list_overdue_allocations(self) -> list[dict]:
        """Overdue week allocations of promised KPs (readable after settle)."""
        return self.repository.list_overdue_allocations()

    def record_plan_exclusions(
        self,
        *,
        plan_id: str,
        exclusions: Sequence[Mapping[str, Any] | Any],
        excluded_by: str,
        user: Mapping[str, Any] | None = None,
        _external_conn: Any | None = None,
    ) -> list[PlanExclusionRecord]:
        """Journal promised-KP exclusions + notify the owner. Does not block build."""
        actor = excluded_by or (_actor(dict(user)) if user else "")
        items = _normalized_exclusions(exclusions)
        if not items:
            return []
        raw = self.repository.record_exclusions(
            plan_id=str(plan_id),
            items=items,
            excluded_by=actor,
            created_at=self._moment(),
            _external_conn=_external_conn,
        )
        return [
            PlanExclusionRecord(
                exclusion_id=int(row["exclusion_id"]),
                notification_id=row["notification_id"],
                kp_id=int(row["kp_id"]),
                week_start=_as_date(row["week_start"]),
                reason=str(row["reason"]),
            )
            for row in raw
        ]

    def list_plan_exclusions(
        self,
        *,
        plan_id: str | None = None,
        kp_id: int | None = None,
    ) -> list[dict]:
        """Read the exclusion journal written at plan build."""
        return self.repository.list_exclusions(plan_id=plan_id, kp_id=kp_id)

    def release_on_delete(
        self,
        kp_id: int,
        *,
        _external_conn: Any | None = None,
    ) -> tuple[int, ...]:
        """Release active promise and hold so delete frees weekly buckets."""
        return self.repository.release_active_for_kp(
            int(kp_id),
            _external_conn=_external_conn,
        )

    def recalc_on_composition_change(
        self,
        kp_id: int,
        *,
        raw: dict | None = None,
    ) -> PromiseRecalcResult:
        """Rewrite active promise tracks/window; notify only if date moves later."""
        empty = PromiseRecalcResult(
            promise_id=None,
            old_promised_date=None,
            new_promised_date=None,
            tracks=0,
            notified=False,
            notification_id=None,
        )
        current = self.repository.get_active_promise(int(kp_id))
        if current is None:
            return empty
        offer = raw if raw is not None else self._load_kp(int(kp_id))
        quote = self._compute_quote(offer, exclude_kp_id=int(kp_id))
        old_date = _as_date(current["promised_date"])
        promise_id = int(current["id"])
        if quote.tracks <= 0 or quote.window is None:
            self.repository.release_active_for_kp(int(kp_id), kinds=("promise",))
            return PromiseRecalcResult(
                promise_id=promise_id,
                old_promised_date=old_date,
                new_promised_date=None,
                tracks=0,
                notified=False,
                notification_id=None,
            )
        new_date = quote.window.promised_date
        notify = new_date > old_date
        payload = None
        kind = None
        if notify:
            kind = KIND_PROMISED_DATE_SHIFTED
            payload = {
                "kp_id": int(kp_id),
                "old_promised_date": old_date.isoformat(),
                "new_promised_date": new_date.isoformat(),
                "tracks": quote.tracks,
            }
        notification_id = self.repository.apply_promise_recalc(
            promise_id=promise_id,
            kp_id=int(kp_id),
            tracks_total=quote.tracks,
            promised_date=new_date,
            allocations=quote.window.allocations,
            created_at=self._moment(),
            notify_kind=kind,
            notify_payload=payload,
        )
        return PromiseRecalcResult(
            promise_id=promise_id,
            old_promised_date=old_date,
            new_promised_date=new_date,
            tracks=quote.tracks,
            notified=notify and notification_id is not None,
            notification_id=notification_id,
        )

    def _compute_quote(self, raw: dict, *, exclude_kp_id: int | None) -> PromiseQuote:
        today = self._today or self._moment().date()
        occupancy = self._occupancy()
        knob = self.repository.get_promise_tracks_per_day()
        buffer = self.repository.get_promise_buffer()
        promised = self.repository.sum_promised_by_week(exclude_kp_id=exclude_kp_id)
        held = self.repository.sum_held_by_week(
            today=today,
            now=self._moment(),
            exclude_kp_id=exclude_kp_id,
        )
        is_workday = self._resolve_workday()
        weeks = build_weeks(
            today,
            occupancy,
            promised_by_week=promised,
            held_by_week=held,
            knob=knob,
            week_count=self._week_count,
            is_workday=is_workday,
        )
        return build_quote(
            _total_length_m(raw),
            weeks,
            today=today,
            knob=knob,
            buffer=buffer,
            is_workday=is_workday,
        )

    def _moment(self) -> datetime:
        if self._now is not None:
            return self._now
        if self._today is not None:
            return datetime.combine(self._today, time(12, 0, 0))
        return datetime.now()

    def _load_kp(self, kp_id: int) -> dict:
        if self._kp_loader is not None:
            raw = self._kp_loader(kp_id)
        else:
            raw = KpArchiveRepository(self.db_path).get_by_id(kp_id)
        if not raw:
            raise PromiseNotFoundError(f"КП №{kp_id} не найдено")
        return raw

    def _occupancy(self) -> Mapping[date | str, int]:
        if self._occupancy_loader is None:
            return load_plan_occupancy(db_path=self.db_path)
        try:
            occupancy = self._occupancy_loader()
        except OccupancyUnavailableError:
            raise
        except Exception as exc:
            logger.exception("promise quote: occupancy loader failed")
            raise OccupancyUnavailableError(_OCCUPANCY_UNAVAILABLE) from exc
        if occupancy is None or not isinstance(occupancy, Mapping):
            raise OccupancyUnavailableError(_OCCUPANCY_UNAVAILABLE)
        return occupancy

    def _resolve_workday(self) -> WorkdayFn:
        if self._is_workday is not None:
            return self._is_workday
        return workday_predicate(load_holidays(), load_extra_workdays())
