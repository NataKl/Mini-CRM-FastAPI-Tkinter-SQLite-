"""
crm_api.py
──────────
FastAPI-бэкэнд мини-CRM.

Запуск:
    venv\\Scripts\\python.exe crm.py
    # или напрямую через uvicorn:
    venv\\Scripts\\uvicorn.exe crm_api:app --reload

Документация API (автогенерация):
    http://127.0.0.1:8000/docs   ← Swagger UI
    http://127.0.0.1:8000/redoc  ← ReDoc
"""

import os
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Query, status
from fastapi.responses import JSONResponse

from crm_db import CRMDatabase
from crm_models import (
    ClientCreate, ClientUpdate, ClientResponse, ClientStatus,
    DealCreate,   DealUpdate,   DealResponse,   DealStatus,
    TaskCreate,   TaskUpdate,   TaskResponse,   TaskStatus, TaskPriority,
)

# ── инициализация ─────────────────────────────────────────────────────────────

DB_PATH = os.getenv("CRM_DB_PATH", "crm.db")
db = CRMDatabase(db_path=DB_PATH)

app = FastAPI(
    title="Mini CRM API",
    description=(
        "REST API для управления клиентами, сделками и задачами.\n\n"
        "**Сущности:**\n"
        "- **Clients** — клиенты / лиды с отслеживанием статуса\n"
        "- **Deals** — сделки и заказы (опциональная связь с клиентом)\n"
        "- **Tasks** — задачи и напоминания (связь с клиентом и/или сделкой)\n"
    ),
    version="1.0.0",
)


# ── guard helpers ─────────────────────────────────────────────────────────────

def _get_client_or_404(client_id: int):
    obj = db.get_client(client_id)
    if not obj:
        raise HTTPException(status_code=404, detail=f"Клиент {client_id} не найден")
    return obj


def _get_deal_or_404(deal_id: int):
    obj = db.get_deal(deal_id)
    if not obj:
        raise HTTPException(status_code=404, detail=f"Сделка {deal_id} не найдена")
    return obj


def _get_task_or_404(task_id: int):
    obj = db.get_task(task_id)
    if not obj:
        raise HTTPException(status_code=404, detail=f"Задача {task_id} не найдена")
    return obj


# ══════════════════════════════════════════════════════════════════════════════
#  DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

@app.get(
    "/dashboard",
    tags=["Dashboard"],
    summary="Сводная статистика CRM",
)
def dashboard():
    """
    Возвращает агрегированную статистику:
    - количество клиентов по статусам
    - количество и суммы сделок
    - количество задач (в т. ч. просроченных)
    """
    return db.get_dashboard_stats()


# ══════════════════════════════════════════════════════════════════════════════
#  CLIENTS
# ══════════════════════════════════════════════════════════════════════════════

@app.get(
    "/clients",
    response_model=List[ClientResponse],
    tags=["Clients"],
    summary="Список клиентов",
)
def list_clients(
    search: Optional[str] = Query(None, description="Поиск по имени, email, компании, телефону"),
    status: Optional[ClientStatus] = Query(None, description="Фильтр по статусу"),
    limit:  int = Query(100, ge=1, le=1000, description="Записей на страницу"),
    offset: int = Query(0,   ge=0,          description="Смещение для пагинации"),
):
    return db.list_clients(search=search, status=status, limit=limit, offset=offset)


@app.get(
    "/clients/{client_id}",
    response_model=ClientResponse,
    tags=["Clients"],
    summary="Получить клиента по ID",
)
def get_client(client_id: int):
    return _get_client_or_404(client_id)


@app.post(
    "/clients",
    response_model=ClientResponse,
    status_code=status.HTTP_201_CREATED,
    tags=["Clients"],
    summary="Создать клиента",
)
def create_client(data: ClientCreate):
    try:
        return db.create_client(data)
    except Exception as exc:
        if "UNIQUE" in str(exc):
            raise HTTPException(status_code=409, detail="Клиент с таким email уже существует")
        raise HTTPException(status_code=500, detail=str(exc))


@app.put(
    "/clients/{client_id}",
    response_model=ClientResponse,
    tags=["Clients"],
    summary="Обновить клиента (частично)",
)
def update_client(client_id: int, data: ClientUpdate):
    _get_client_or_404(client_id)
    try:
        return db.update_client(client_id, data)
    except Exception as exc:
        if "UNIQUE" in str(exc):
            raise HTTPException(status_code=409, detail="Email уже занят другим клиентом")
        raise HTTPException(status_code=500, detail=str(exc))


@app.delete(
    "/clients/{client_id}",
    status_code=status.HTTP_204_NO_CONTENT,
    tags=["Clients"],
    summary="Удалить клиента",
)
def delete_client(client_id: int):
    _get_client_or_404(client_id)
    db.delete_client(client_id)


@app.patch(
    "/clients/{client_id}/archive",
    response_model=ClientResponse,
    tags=["Clients"],
    summary="Архивировать клиента",
    description="Мягкое удаление: устанавливает статус `archived`, не стирая данные.",
)
def archive_client(client_id: int):
    _get_client_or_404(client_id)
    return db.archive_client(client_id)


@app.patch(
    "/clients/{client_id}/status",
    response_model=ClientResponse,
    tags=["Clients"],
    summary="Изменить статус клиента",
)
def set_client_status(
    client_id: int,
    status: ClientStatus = Query(..., description="Новый статус"),
):
    _get_client_or_404(client_id)
    return db.update_client(client_id, ClientUpdate(status=status))


@app.get(
    "/clients/{client_id}/deals",
    response_model=List[DealResponse],
    tags=["Clients"],
    summary="Все сделки клиента",
)
def get_client_deals(client_id: int):
    _get_client_or_404(client_id)
    return db.get_client_deals(client_id)


@app.get(
    "/clients/{client_id}/tasks",
    response_model=List[TaskResponse],
    tags=["Clients"],
    summary="Все задачи клиента",
)
def get_client_tasks(client_id: int):
    _get_client_or_404(client_id)
    return db.get_client_tasks(client_id)


# ══════════════════════════════════════════════════════════════════════════════
#  DEALS
# ══════════════════════════════════════════════════════════════════════════════

@app.get(
    "/deals",
    response_model=List[DealResponse],
    tags=["Deals"],
    summary="Список сделок",
)
def list_deals(
    search:    Optional[str]        = Query(None, description="Поиск по названию и описанию"),
    status:    Optional[DealStatus] = Query(None, description="Фильтр по статусу"),
    client_id: Optional[int]        = Query(None, description="Фильтр по клиенту"),
    limit:     int = Query(100, ge=1, le=1000),
    offset:    int = Query(0,   ge=0),
):
    return db.list_deals(
        search=search, status=status, client_id=client_id,
        limit=limit, offset=offset,
    )


@app.get(
    "/deals/{deal_id}",
    response_model=DealResponse,
    tags=["Deals"],
    summary="Получить сделку по ID",
)
def get_deal(deal_id: int):
    return _get_deal_or_404(deal_id)


@app.post(
    "/deals",
    response_model=DealResponse,
    status_code=status.HTTP_201_CREATED,
    tags=["Deals"],
    summary="Создать сделку",
)
def create_deal(data: DealCreate):
    if data.client_id is not None:
        _get_client_or_404(data.client_id)
    return db.create_deal(data)


@app.put(
    "/deals/{deal_id}",
    response_model=DealResponse,
    tags=["Deals"],
    summary="Обновить сделку (частично)",
)
def update_deal(deal_id: int, data: DealUpdate):
    _get_deal_or_404(deal_id)
    if data.client_id is not None:
        _get_client_or_404(data.client_id)
    return db.update_deal(deal_id, data)


@app.delete(
    "/deals/{deal_id}",
    status_code=status.HTTP_204_NO_CONTENT,
    tags=["Deals"],
    summary="Удалить сделку",
)
def delete_deal(deal_id: int):
    _get_deal_or_404(deal_id)
    db.delete_deal(deal_id)


@app.patch(
    "/deals/{deal_id}/status",
    response_model=DealResponse,
    tags=["Deals"],
    summary="Изменить статус сделки",
)
def update_deal_status(
    deal_id: int,
    status: DealStatus = Query(..., description="Новый статус сделки"),
):
    _get_deal_or_404(deal_id)
    return db.update_deal_status(deal_id, status)


@app.patch(
    "/deals/{deal_id}/attach-client",
    response_model=DealResponse,
    tags=["Deals"],
    summary="Привязать/отвязать клиента к сделке",
)
def attach_client_to_deal(
    deal_id:   int,
    client_id: Optional[int] = Query(None, description="ID клиента; null — отвязать"),
):
    _get_deal_or_404(deal_id)
    if client_id is not None:
        _get_client_or_404(client_id)
    return db.update_deal(deal_id, DealUpdate(client_id=client_id))


@app.get(
    "/deals/{deal_id}/tasks",
    response_model=List[TaskResponse],
    tags=["Deals"],
    summary="Все задачи по сделке",
)
def get_deal_tasks(deal_id: int):
    _get_deal_or_404(deal_id)
    return db.get_deal_tasks(deal_id)


# ══════════════════════════════════════════════════════════════════════════════
#  TASKS
# ══════════════════════════════════════════════════════════════════════════════

@app.get(
    "/tasks",
    response_model=List[TaskResponse],
    tags=["Tasks"],
    summary="Список задач",
)
def list_tasks(
    search:       Optional[str]          = Query(None, description="Поиск по заголовку и описанию"),
    status:       Optional[TaskStatus]   = Query(None, description="Фильтр по статусу"),
    priority:     Optional[TaskPriority] = Query(None, description="Фильтр по приоритету"),
    client_id:    Optional[int]          = Query(None, description="Фильтр по клиенту"),
    deal_id:      Optional[int]          = Query(None, description="Фильтр по сделке"),
    overdue_only: bool                   = Query(False, description="Только просроченные"),
    limit:        int = Query(100, ge=1, le=1000),
    offset:       int = Query(0,   ge=0),
):
    return db.list_tasks(
        search=search, status=status, priority=priority,
        client_id=client_id, deal_id=deal_id,
        overdue_only=overdue_only, limit=limit, offset=offset,
    )


@app.get(
    "/tasks/{task_id}",
    response_model=TaskResponse,
    tags=["Tasks"],
    summary="Получить задачу по ID",
)
def get_task(task_id: int):
    return _get_task_or_404(task_id)


@app.post(
    "/tasks",
    response_model=TaskResponse,
    status_code=status.HTTP_201_CREATED,
    tags=["Tasks"],
    summary="Создать задачу / напоминание",
)
def create_task(data: TaskCreate):
    if data.client_id is not None:
        _get_client_or_404(data.client_id)
    if data.deal_id is not None:
        _get_deal_or_404(data.deal_id)
    return db.create_task(data)


@app.put(
    "/tasks/{task_id}",
    response_model=TaskResponse,
    tags=["Tasks"],
    summary="Обновить задачу (частично)",
)
def update_task(task_id: int, data: TaskUpdate):
    _get_task_or_404(task_id)
    if data.client_id is not None:
        _get_client_or_404(data.client_id)
    if data.deal_id is not None:
        _get_deal_or_404(data.deal_id)
    return db.update_task(task_id, data)


@app.delete(
    "/tasks/{task_id}",
    status_code=status.HTTP_204_NO_CONTENT,
    tags=["Tasks"],
    summary="Удалить задачу",
)
def delete_task(task_id: int):
    _get_task_or_404(task_id)
    db.delete_task(task_id)


@app.patch(
    "/tasks/{task_id}/complete",
    response_model=TaskResponse,
    tags=["Tasks"],
    summary="Отметить задачу как выполненную",
)
def complete_task(task_id: int):
    _get_task_or_404(task_id)
    return db.complete_task(task_id)


@app.patch(
    "/tasks/{task_id}/status",
    response_model=TaskResponse,
    tags=["Tasks"],
    summary="Изменить статус задачи",
)
def set_task_status(
    task_id: int,
    status: TaskStatus = Query(..., description="Новый статус задачи"),
):
    _get_task_or_404(task_id)
    return db.update_task(task_id, TaskUpdate(status=status))
