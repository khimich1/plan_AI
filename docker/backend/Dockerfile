# syntax=docker/dockerfile:1
# Контекст сборки — корень репозитория:
#   docker build -f docker/backend/Dockerfile -t plan-wed-backend:latest .
#
# Статический SPA можно отдавать отдельным образом frontend (см. frontend/Dockerfile).

FROM python:3.12-slim-bookworm AS runtime

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        ca-certificates \
        coinor-cbc \
        fonts-dejavu-core \
        gosu \
    && rm -rf /var/lib/apt/lists/* \
    && adduser --disabled-password --gecos '' appuser

COPY docker/entrypoint.sh /entrypoint.sh
RUN chmod 0755 /entrypoint.sh

COPY requirements.txt .
RUN python -m pip install --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY --chown=appuser:appuser . .

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=40s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=4)"

ENTRYPOINT ["/entrypoint.sh"]
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
