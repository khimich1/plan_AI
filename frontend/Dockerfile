# syntax=docker/dockerfile:1
# Сборка и отдача Vite SPA (base=/commercial-offer/) через Nginx с проксированием API на backend.
#
# Контекст — каталог frontend/:
#   docker build -f Dockerfile -t plan-wed-frontend:latest .
#
# Сборочные аргументы (необязательно; по умолчанию API берётся с того же origin, что открыл браузер):
#   docker build --build-arg VITE_API_BASE_URL=http://zavodstart.ru -t plan-wed-frontend:latest .

FROM node:22-bookworm-slim AS build

WORKDIR /src

ARG VITE_API_BASE_URL=""
ENV VITE_API_BASE_URL=${VITE_API_BASE_URL}

COPY package.json package-lock.json ./
RUN npm ci

COPY . .
RUN npm run build

# --- Production: nginx ---
FROM nginx:1.26-alpine

COPY nginx.docker.conf /etc/nginx/conf.d/default.conf
COPY --from=build /src/dist/ /usr/share/nginx/html/commercial-offer/

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=5s --start-period=5s --retries=3 \
    CMD wget -qO- http://127.0.0.1/health >/dev/null 2>&1 || exit 1
