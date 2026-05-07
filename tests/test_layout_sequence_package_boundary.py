# -*- coding: utf-8 -*-
"""Guard against regressing PLATE-LAYOUT-PKG: module `layout_sequence.py` shadowing the package."""

from __future__ import annotations

import pathlib


def test_viz_layout_sequence_is_package_not_top_level_py() -> None:
    """A sibling `viz_modules/layout_sequence.py` preempts `viz_modules/layout_sequence/` on import."""
    import viz_modules.layout_sequence as ls

    assert getattr(ls, "__path__", None) is not None
    bf = pathlib.Path(ls.__file__).resolve()
    assert bf.name == "__init__.py", f"expected package __init__, got {bf}"
