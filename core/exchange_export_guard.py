"""Refuse production startup when the 1C exchange directory is not mounted."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from pathlib import Path


class ExchangeExportMisconfigured(RuntimeError):
    """Production is about to write invoices outside the directory 1C reads."""


def fail_fast_if_exchange_dir_not_mounted(
    app_env: str,
    directory: Path,
    *,
    environ: Mapping[str, str],
    mount_points: Iterable[Path],
) -> None:
    """In production, the export directory must be its own mount point.

    A path that only sits inside the database volume is not the SFTP folder.
    Development keeps the in-project fallback and is not checked.
    """
    if app_env.strip().lower() != "production":
        return
    if "EXCHANGE_EXPORT_DIR" not in environ:
        raise ExchangeExportMisconfigured("EXCHANGE_EXPORT_DIR не задан")
    mounted = {path.resolve() for path in mount_points}
    if directory.resolve() not in mounted:
        raise ExchangeExportMisconfigured("каталог обмена не смонтирован отдельно")


def read_mount_points() -> tuple[Path, ...]:
    """Mount points of this process, from ``/proc/self/mountinfo``."""
    text = Path("/proc/self/mountinfo").read_text(encoding="utf-8")
    points: list[Path] = []
    for line in text.splitlines():
        fields = line.split()
        if len(fields) < 5:
            continue
        points.append(Path(_unescape_mount_field(fields[4])))
    return tuple(points)


def _unescape_mount_field(value: str) -> str:
    """Kernel octal escapes: space is ``\\040``, backslash is ``\\134``."""
    out: list[str] = []
    index = 0
    while index < len(value):
        if (
            value[index] == "\\"
            and index + 3 < len(value)
            and value[index + 1 : index + 4].isdigit()
        ):
            out.append(chr(int(value[index + 1 : index + 4], 8)))
            index += 4
            continue
        out.append(value[index])
        index += 1
    return "".join(out)
