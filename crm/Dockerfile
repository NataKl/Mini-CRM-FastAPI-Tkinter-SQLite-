# ─────────────────────────────────────────────────────────────────────────────
# Mini CRM — FastAPI backend
#
# Образ создаётся с зависимостями, но БЕЗ исходного кода.
# Код монтируется через bind-mount в docker-compose.yml (./crm → /app),
# что позволяет редактировать файлы без пересборки образа.
#
# Hot-reload: uvicorn --reload + watchfiles отслеживает /app на изменения.
# При сохранении любого .py файла сервер перезапускается автоматически.
# ─────────────────────────────────────────────────────────────────────────────

FROM python:3.11-slim

# Метаданные
LABEL maintainer="mini-crm"
LABEL description="Mini CRM FastAPI backend with hot-reload"

# Переменные сборки
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /app

# ── Зависимости ───────────────────────────────────────────────────────────────
# requirements.txt берётся из корня проекта (build context = .)
COPY requirements.txt /tmp/requirements.txt

RUN pip install \
        -r /tmp/requirements.txt \
        "uvicorn[standard]" \
        watchfiles \
    && rm /tmp/requirements.txt

# ── Директория для SQLite ─────────────────────────────────────────────────────
# Данные хранятся в /data (именованный volume из compose),
# чтобы база пережила пересоздание контейнера.
RUN mkdir -p /data

EXPOSE 8000

# ── Точка входа с hot-reload ──────────────────────────────────────────────────
# --reload-dir /app  → watchfiles следит за всем кодом в /app
# --reload-delay 0.25 → небольшая задержка, чтобы избежать двойного рестарта
#                       при одновременном сохранении нескольких файлов
CMD ["uvicorn", "crm_api:app", \
     "--host", "0.0.0.0", \
     "--port", "8000", \
     "--reload", \
     "--reload-dir", "/app", \
     "--reload-delay", "0.25", \
     "--log-level", "info"]
