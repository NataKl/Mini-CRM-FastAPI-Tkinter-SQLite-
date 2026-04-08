"""
crm_models.py
─────────────
Два слоя определений:

1. SQL_CREATE_* — строки CREATE TABLE для SQLite3 (используются в crm_db.py).
2. Pydantic-схемы — валидация входящих данных и сериализация ответов (используются в crm_api.py).

Таблицы:
  clients — клиенты / лиды
  deals   — сделки / заказы (опционально привязаны к клиенту)
  tasks   — задачи / напоминания (опционально привязаны к клиенту и/или сделке)
"""

from enum import Enum
from typing import Optional

from pydantic import BaseModel, EmailStr, Field


# ══════════════════════════════════════════════════════════════════════════════
#  SQLite DDL
# ══════════════════════════════════════════════════════════════════════════════

SQL_CREATE_CLIENTS = """
CREATE TABLE IF NOT EXISTS clients (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT    NOT NULL,
    email       TEXT    UNIQUE,
    phone       TEXT,
    company     TEXT,
    source      TEXT,
    status      TEXT    NOT NULL DEFAULT 'lead',
    notes       TEXT,
    created_at  TEXT    NOT NULL,
    updated_at  TEXT    NOT NULL
)
"""

SQL_CREATE_DEALS = """
CREATE TABLE IF NOT EXISTS deals (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    title       TEXT    NOT NULL,
    client_id   INTEGER REFERENCES clients(id) ON DELETE SET NULL,
    amount      REAL    NOT NULL DEFAULT 0.0,
    currency    TEXT    NOT NULL DEFAULT 'RUB',
    status      TEXT    NOT NULL DEFAULT 'new',
    description TEXT,
    due_date    TEXT,
    created_at  TEXT    NOT NULL,
    updated_at  TEXT    NOT NULL
)
"""

SQL_CREATE_TASKS = """
CREATE TABLE IF NOT EXISTS tasks (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    title       TEXT    NOT NULL,
    description TEXT,
    client_id   INTEGER REFERENCES clients(id) ON DELETE SET NULL,
    deal_id     INTEGER REFERENCES deals(id)   ON DELETE SET NULL,
    status      TEXT    NOT NULL DEFAULT 'pending',
    priority    TEXT    NOT NULL DEFAULT 'medium',
    due_date    TEXT,
    created_at  TEXT    NOT NULL,
    updated_at  TEXT    NOT NULL
)
"""

# Индексы для ускорения фильтрации и поиска
SQL_CREATE_INDEXES = [
    "CREATE INDEX IF NOT EXISTS idx_clients_name    ON clients(name)",
    "CREATE INDEX IF NOT EXISTS idx_clients_status  ON clients(status)",
    "CREATE INDEX IF NOT EXISTS idx_deals_client    ON deals(client_id)",
    "CREATE INDEX IF NOT EXISTS idx_deals_status    ON deals(status)",
    "CREATE INDEX IF NOT EXISTS idx_tasks_client    ON tasks(client_id)",
    "CREATE INDEX IF NOT EXISTS idx_tasks_deal      ON tasks(deal_id)",
    "CREATE INDEX IF NOT EXISTS idx_tasks_status    ON tasks(status)",
    "CREATE INDEX IF NOT EXISTS idx_tasks_priority  ON tasks(priority)",
    "CREATE INDEX IF NOT EXISTS idx_tasks_due_date  ON tasks(due_date)",
]


# ══════════════════════════════════════════════════════════════════════════════
#  Enums
# ══════════════════════════════════════════════════════════════════════════════

class ClientStatus(str, Enum):
    lead     = "lead"       # потенциальный клиент
    active   = "active"     # активный клиент
    vip      = "vip"        # VIP-клиент
    inactive = "inactive"   # неактивный
    archived = "archived"   # архивирован


class DealStatus(str, Enum):
    new         = "new"
    in_progress = "in_progress"
    won         = "won"
    lost        = "lost"
    on_hold     = "on_hold"


class TaskStatus(str, Enum):
    pending   = "pending"
    done      = "done"
    cancelled = "cancelled"


class TaskPriority(str, Enum):
    low    = "low"
    medium = "medium"
    high   = "high"


# ══════════════════════════════════════════════════════════════════════════════
#  Client schemas
# ══════════════════════════════════════════════════════════════════════════════

class ClientCreate(BaseModel):
    name:    str            = Field(..., min_length=1, max_length=200)
    email:   Optional[str] = Field(None, max_length=200)
    phone:   Optional[str] = Field(None, max_length=50)
    company: Optional[str] = Field(None, max_length=200)
    source:  Optional[str] = Field(None, max_length=100, description="Откуда пришёл: реклама, звонок, реферал …")
    status:  ClientStatus  = ClientStatus.lead
    notes:   Optional[str] = None


class ClientUpdate(BaseModel):
    name:    Optional[str]          = Field(None, min_length=1, max_length=200)
    email:   Optional[str]          = Field(None, max_length=200)
    phone:   Optional[str]          = Field(None, max_length=50)
    company: Optional[str]          = Field(None, max_length=200)
    source:  Optional[str]          = Field(None, max_length=100)
    status:  Optional[ClientStatus] = None
    notes:   Optional[str]          = None


class ClientResponse(BaseModel):
    id:         int
    name:       str
    email:      Optional[str]
    phone:      Optional[str]
    company:    Optional[str]
    source:     Optional[str]
    status:     str
    notes:      Optional[str]
    created_at: str
    updated_at: str


# ══════════════════════════════════════════════════════════════════════════════
#  Deal schemas
# ══════════════════════════════════════════════════════════════════════════════

class DealCreate(BaseModel):
    title:       str           = Field(..., min_length=1, max_length=300)
    client_id:   Optional[int] = Field(None, description="ID клиента (необязательно)")
    amount:      float         = Field(0.0, ge=0)
    currency:    str           = Field("RUB", max_length=10)
    status:      DealStatus    = DealStatus.new
    description: Optional[str] = None
    due_date:    Optional[str] = Field(None, description="Дата закрытия в формате YYYY-MM-DD")


class DealUpdate(BaseModel):
    title:       Optional[str]        = Field(None, min_length=1, max_length=300)
    client_id:   Optional[int]        = None
    amount:      Optional[float]      = Field(None, ge=0)
    currency:    Optional[str]        = Field(None, max_length=10)
    status:      Optional[DealStatus] = None
    description: Optional[str]        = None
    due_date:    Optional[str]        = None


class DealResponse(BaseModel):
    id:          int
    title:       str
    client_id:   Optional[int]
    amount:      float
    currency:    str
    status:      str
    description: Optional[str]
    due_date:    Optional[str]
    created_at:  str
    updated_at:  str


# ══════════════════════════════════════════════════════════════════════════════
#  Task schemas
# ══════════════════════════════════════════════════════════════════════════════

class TaskCreate(BaseModel):
    title:       str              = Field(..., min_length=1, max_length=300)
    description: Optional[str]   = None
    client_id:   Optional[int]   = Field(None, description="Привязать к клиенту")
    deal_id:     Optional[int]   = Field(None, description="Привязать к сделке")
    status:      TaskStatus      = TaskStatus.pending
    priority:    TaskPriority    = TaskPriority.medium
    due_date:    Optional[str]   = Field(None, description="Срок выполнения YYYY-MM-DD")


class TaskUpdate(BaseModel):
    title:       Optional[str]           = Field(None, min_length=1, max_length=300)
    description: Optional[str]          = None
    client_id:   Optional[int]          = None
    deal_id:     Optional[int]          = None
    status:      Optional[TaskStatus]   = None
    priority:    Optional[TaskPriority] = None
    due_date:    Optional[str]          = None


class TaskResponse(BaseModel):
    id:          int
    title:       str
    description: Optional[str]
    client_id:   Optional[int]
    deal_id:     Optional[int]
    status:      str
    priority:    str
    due_date:    Optional[str]
    created_at:  str
    updated_at:  str
