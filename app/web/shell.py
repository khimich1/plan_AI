from __future__ import annotations

from html import escape
from pathlib import Path

from fastapi.responses import HTMLResponse

from app.core.settings import get_settings


def page(title: str, body: str) -> HTMLResponse:
    html = f"""
    <!doctype html>
    <html lang="ru">
      <head>
        <meta charset="utf-8">
        <title>{escape(title)}</title>
        <style>
          body {{ font-family: Arial, sans-serif; margin: 2rem; }}
          nav a {{ margin-right: 1rem; }}
          table {{ border-collapse: collapse; width: 100%; margin-top: 1rem; }}
          th, td {{ border: 1px solid #ccc; padding: 0.5rem; text-align: left; }}
          .card {{ border: 1px solid #ddd; padding: 1rem; margin-bottom: 1rem; border-radius: 8px; }}
        </style>
      </head>
      <body>{body}</body>
    </html>
    """
    return HTMLResponse(html)


def render_frontend_shell() -> HTMLResponse:
    settings = get_settings()
    index_path = settings.frontend_dist_dir / "index.html"
    if not index_path.exists():
        body = """
        <h1>Frontend build не найден</h1>
        <p>Соберите React-приложение в директорию <code>frontend/dist</code>, затем откройте страницу снова.</p>
        <p>Для локальной разработки используйте Vite dev server из директории <code>frontend/</code>.</p>
        """
        return page("Frontend build missing", body)
    return HTMLResponse(index_path.read_text(encoding="utf-8"))


def resolve_frontend_asset(asset_path: str) -> Path | None:
    settings = get_settings()
    assets_dir = (settings.frontend_dist_dir / "assets").resolve()
    candidate = (assets_dir / asset_path).resolve()
    if assets_dir not in candidate.parents or not candidate.exists():
        return None
    return candidate
