"""One-off: replace gated debug open().write() with write_agent_debug in bot handlers."""

from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

FILES = [
    ROOT / "bot/handlers/production_completion.py",
    ROOT / "bot/handlers/production_execution.py",
    ROOT / "bot/handlers/commercial.py",
    ROOT / "bot/handlers/production_export.py",
    ROOT / "bot/handlers/production_day_view.py",
]

IMPORT_SNIPPET = (
    "from .debug_util import write_agent_debug, write_agent_debug_session"
)

TRY_BLOCK = re.compile(
    r"([ \t]*)try:\s*\n"
    r"[ \t]*with open\(([^,]+), [^)]+\) as \w+:\s*\n"
    r"[ \t]*\w+\.write\(([^\n]+)\)\s*\n"
    r"[ \t]*except Exception:\s*\n"
    r"[ \t]*pass",
    re.MULTILINE,
)

PLAIN_BLOCK = re.compile(
    r"([ \t]*)with open\(([^,]+), [^)]+\) as \w+:\s*\n"
    r"[ \t]*\w+\.write\(([^\n]+)\)",
    re.MULTILINE,
)


def _extract_payload(write_expr: str) -> str | None:
    m = re.search(r"\.dumps\((\{.*\}),\s*ensure_ascii=False\)", write_expr)
    if m:
        return m.group(1)
    m = re.search(r"\.dumps\((\{.*\})\s*\+\s*[\"']\\n[\"']\)", write_expr)
    if m:
        return m.group(1)
    m = re.search(r"\.dumps\((_payload)\s*\+\s*[\"']\\n[\"']\)", write_expr)
    if m:
        return m.group(1)
    return None


def _replace_write(indent: str, log: str, write_expr: str) -> str | None:
    payload = _extract_payload(write_expr)
    if payload is None:
        return None
    return f"{indent}write_agent_debug({log}, {payload})"


def fix_file(path: Path) -> tuple[int, int]:
    text = path.read_text(encoding="utf-8")
    if IMPORT_SNIPPET not in text and "debug_util" not in text:
        if "from core.debug_paths import get_debug_log_path" in text:
            text = text.replace(
                "from core.debug_paths import get_debug_log_path",
                f"from core.debug_paths import get_debug_log_path\n\n{IMPORT_SNIPPET}",
            )
        elif "from core.debug_paths import" in text:
            pass

    n_try = 0

    def try_sub(m: re.Match) -> str:
        nonlocal n_try
        repl = _replace_write(m.group(1), m.group(2), m.group(3))
        if repl is None:
            return m.group(0)
        n_try += 1
        return f"{m.group(1)}try:\n{m.group(1)}    {repl.strip()}\n{m.group(1)}except Exception:\n{m.group(1)}    pass"

    text = TRY_BLOCK.sub(try_sub, text)

    n_plain = 0

    def plain_sub(m: re.Match) -> str:
        nonlocal n_plain
        repl = _replace_write(m.group(1), m.group(2), m.group(3))
        if repl is None:
            return m.group(0)
        n_plain += 1
        return repl

    text = PLAIN_BLOCK.sub(plain_sub, text)

    # commercial one-liner: open(...).write(...)
    text, n_com = re.subn(
        r"open\(([^,]+), [^)]+\)\.write\(\s*([^)]+)\.dumps\((\{.*?\}),",
        r"write_agent_debug(\1, \3)",
        text,
        count=0,
    )

    path.write_text(text, encoding="utf-8")
    return n_try + n_plain, n_com


def main() -> None:
    total = 0
    for f in FILES:
        if not f.exists():
            continue
        a, b = fix_file(f)
        remaining = len(re.findall(r"with open\([^)]*debug", f.read_text(encoding="utf-8"), re.I))
        remaining += len(re.findall(r"open\(_DEBUG", f.read_text(encoding="utf-8")))
        print(f"{f.name}: replaced={a+b}, remaining_debug_open={remaining}")
        total += a + b
    print("total replaced", total)


if __name__ == "__main__":
    main()
