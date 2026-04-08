"""
crm_db.py
─────────
CRMDatabase — класс с полным набором CRUD-операций поверх SQLite3.

Особенности:
  • Каждый вызов открывает собственное соединение через контекстный менеджер
    (thread-safe, WAL-mode для конкурентного чтения).
  • При инициализации автоматически создаёт все таблицы и индексы, если их нет.
  • Поиск через LIKE по ключевым текстовым полям каждой сущности.
  • Связи: сделка → клиент, задача → клиент и/или сделка.
"""

import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from typing import Any, Dict, Generator, List, Optional

from crm_models import (
    SQL_CREATE_CLIENTS,
    SQL_CREATE_DEALS,
    SQL_CREATE_TASKS,
    SQL_CREATE_INDEXES,
    ClientCreate,
    ClientUpdate,
    DealCreate,
    DealUpdate,
    TaskCreate,
    TaskUpdate,
)


# ── helpers ───────────────────────────────────────────────────────────────────

def _now() -> str:
    """Текущее время UTC в формате ISO-8601."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")


def _row(row: Optional[sqlite3.Row]) -> Optional[Dict[str, Any]]:
    return dict(row) if row else None


def _contains_text(value: Optional[str], needle: str) -> bool:
    """Unicode-safe substring search for UI filters."""
    return needle in (value or "").casefold()


# ══════════════════════════════════════════════════════════════════════════════
#  CRMDatabase
# ══════════════════════════════════════════════════════════════════════════════

class CRMDatabase:
    """
    SQLite3-backed CRM data layer.

    Args:
        db_path: Путь к файлу базы данных (создаётся автоматически).
    """

    def __init__(self, db_path: str = "crm.db") -> None:
        self.db_path = db_path
        self._init_db()

    # ── connection ────────────────────────────────────────────────────────────

    @contextmanager
    def _connect(self) -> Generator[sqlite3.Connection, None, None]:
        """
        Контекстный менеджер соединения.
        Автоматически коммитит при успехе, откатывает при исключении.
        """
        conn = sqlite3.connect(self.db_path, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode = WAL")   # concurrent reads
        conn.execute("PRAGMA foreign_keys  = ON")   # enforce FK constraints
        conn.execute("PRAGMA busy_timeout  = 5000")  # ms to wait on lock
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    # ── init ──────────────────────────────────────────────────────────────────

    def _init_db(self) -> None:
        """Создаёт таблицы и индексы, если они ещё не существуют."""
        with self._connect() as conn:
            conn.execute(SQL_CREATE_CLIENTS)
            conn.execute(SQL_CREATE_DEALS)
            conn.execute(SQL_CREATE_TASKS)
            for sql in SQL_CREATE_INDEXES:
                conn.execute(sql)

    # ══════════════════════════════════════════════════════════════════════════
    #  CLIENTS
    # ══════════════════════════════════════════════════════════════════════════

    def create_client(self, data: ClientCreate) -> Dict[str, Any]:
        now = _now()
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO clients
                    (name, email, phone, company, source, status, notes, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    data.name, data.email, data.phone, data.company,
                    data.source, data.status, data.notes, now, now,
                ),
            )
            row = conn.execute(
                "SELECT * FROM clients WHERE id = ?", (cur.lastrowid,)
            ).fetchone()
            return dict(row)

    def get_client(self, client_id: int) -> Optional[Dict[str, Any]]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM clients WHERE id = ?", (client_id,)
            ).fetchone()
            return _row(row)

    def list_clients(
        self,
        search:  Optional[str] = None,
        status:  Optional[str] = None,
        limit:   int = 100,
        offset:  int = 0,
    ) -> List[Dict[str, Any]]:
        """
        Список клиентов с опциональной фильтрацией.
        Поиск выполняется только по полям, которые видны в таблице UI:
        name, email, company, phone.
        """
        conditions: List[str] = []
        params: List[Any] = []

        if status:
            conditions.append("status = ?")
            params.append(status)

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT * FROM clients {where} ORDER BY updated_at DESC",
                params,
            ).fetchall()
            items = [dict(r) for r in rows]

        if search:
            needle = search.casefold()
            items = [
                item for item in items
                if (
                    _contains_text(item.get("name"), needle)
                    or _contains_text(item.get("email"), needle)
                    or _contains_text(item.get("company"), needle)
                    or _contains_text(item.get("phone"), needle)
                )
            ]

        return items[offset: offset + limit]

    def update_client(
        self, client_id: int, data: ClientUpdate
    ) -> Optional[Dict[str, Any]]:
        """
        Частичное обновление — меняет только явно переданные поля.
        Поля, не включённые в запрос, остаются без изменений.
        """
        updates = data.model_dump(exclude_unset=True)
        if not updates:
            return self.get_client(client_id)

        updates["updated_at"] = _now()
        fields = ", ".join(f"{k} = ?" for k in updates)
        values = list(updates.values()) + [client_id]

        with self._connect() as conn:
            conn.execute(
                f"UPDATE clients SET {fields} WHERE id = ?", values
            )
            row = conn.execute(
                "SELECT * FROM clients WHERE id = ?", (client_id,)
            ).fetchone()
            return _row(row)

    def delete_client(self, client_id: int) -> bool:
        """Удаляет клиента навсегда. Возвращает True если запись существовала."""
        with self._connect() as conn:
            cur = conn.execute("DELETE FROM clients WHERE id = ?", (client_id,))
            return cur.rowcount > 0

    def archive_client(self, client_id: int) -> Optional[Dict[str, Any]]:
        """Переводит статус клиента в 'archived' (мягкое удаление)."""
        with self._connect() as conn:
            conn.execute(
                "UPDATE clients SET status = 'archived', updated_at = ? WHERE id = ?",
                (_now(), client_id),
            )
            row = conn.execute(
                "SELECT * FROM clients WHERE id = ?", (client_id,)
            ).fetchone()
            return _row(row)

    def get_client_deals(self, client_id: int) -> List[Dict[str, Any]]:
        """Все сделки конкретного клиента."""
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM deals WHERE client_id = ? ORDER BY updated_at DESC",
                (client_id,),
            ).fetchall()
            return [dict(r) for r in rows]

    def get_client_tasks(self, client_id: int) -> List[Dict[str, Any]]:
        """Все задачи конкретного клиента."""
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM tasks WHERE client_id = ? ORDER BY priority DESC, due_date ASC",
                (client_id,),
            ).fetchall()
            return [dict(r) for r in rows]

    # ══════════════════════════════════════════════════════════════════════════
    #  DEALS
    # ══════════════════════════════════════════════════════════════════════════

    def create_deal(self, data: DealCreate) -> Dict[str, Any]:
        now = _now()
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO deals
                    (title, client_id, amount, currency, status, description, due_date, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    data.title, data.client_id, data.amount, data.currency,
                    data.status, data.description, data.due_date, now, now,
                ),
            )
            row = conn.execute(
                "SELECT * FROM deals WHERE id = ?", (cur.lastrowid,)
            ).fetchone()
            return dict(row)

    def get_deal(self, deal_id: int) -> Optional[Dict[str, Any]]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM deals WHERE id = ?", (deal_id,)
            ).fetchone()
            return _row(row)

    def list_deals(
        self,
        search:    Optional[str] = None,
        status:    Optional[str] = None,
        client_id: Optional[int] = None,
        limit:     int = 100,
        offset:    int = 0,
    ) -> List[Dict[str, Any]]:
        """
        Список сделок с опциональной фильтрацией.
        Поиск выполняется только по названию сделки, потому что именно оно
        отображается в таблице UI и не должно давать «скрытых» совпадений по
        полю description.
        """
        conditions: List[str] = []
        params: List[Any] = []

        if status:
            conditions.append("status = ?")
            params.append(status)
        if client_id is not None:
            conditions.append("client_id = ?")
            params.append(client_id)

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        with self._connect() as conn:
            rows = conn.execute(
                f"SELECT * FROM deals {where} ORDER BY updated_at DESC",
                params,
            ).fetchall()
            items = [dict(r) for r in rows]

        if search:
            needle = search.casefold()
            items = [
                item for item in items
                if _contains_text(item.get("title"), needle)
            ]

        return items[offset: offset + limit]

    def update_deal(
        self, deal_id: int, data: DealUpdate
    ) -> Optional[Dict[str, Any]]:
        updates = data.model_dump(exclude_unset=True)
        if not updates:
            return self.get_deal(deal_id)

        updates["updated_at"] = _now()
        fields = ", ".join(f"{k} = ?" for k in updates)
        values = list(updates.values()) + [deal_id]

        with self._connect() as conn:
            conn.execute(f"UPDATE deals SET {fields} WHERE id = ?", values)
            row = conn.execute(
                "SELECT * FROM deals WHERE id = ?", (deal_id,)
            ).fetchone()
            return _row(row)

    def update_deal_status(
        self, deal_id: int, status: str
    ) -> Optional[Dict[str, Any]]:
        with self._connect() as conn:
            conn.execute(
                "UPDATE deals SET status = ?, updated_at = ? WHERE id = ?",
                (status, _now(), deal_id),
            )
            row = conn.execute(
                "SELECT * FROM deals WHERE id = ?", (deal_id,)
            ).fetchone()
            return _row(row)

    def delete_deal(self, deal_id: int) -> bool:
        with self._connect() as conn:
            cur = conn.execute("DELETE FROM deals WHERE id = ?", (deal_id,))
            return cur.rowcount > 0

    def get_deal_tasks(self, deal_id: int) -> List[Dict[str, Any]]:
        """Все задачи конкретной сделки."""
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM tasks WHERE deal_id = ? ORDER BY priority DESC, due_date ASC",
                (deal_id,),
            ).fetchall()
            return [dict(r) for r in rows]

    # ══════════════════════════════════════════════════════════════════════════
    #  TASKS
    # ══════════════════════════════════════════════════════════════════════════

    def create_task(self, data: TaskCreate) -> Dict[str, Any]:
        now = _now()
        with self._connect() as conn:
            cur = conn.execute(
                """
                INSERT INTO tasks
                    (title, description, client_id, deal_id, status, priority, due_date, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    data.title, data.description, data.client_id, data.deal_id,
                    data.status, data.priority, data.due_date, now, now,
                ),
            )
            row = conn.execute(
                "SELECT * FROM tasks WHERE id = ?", (cur.lastrowid,)
            ).fetchone()
            return dict(row)

    def get_task(self, task_id: int) -> Optional[Dict[str, Any]]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM tasks WHERE id = ?", (task_id,)
            ).fetchone()
            return _row(row)

    def list_tasks(
        self,
        search:       Optional[str] = None,
        status:       Optional[str] = None,
        priority:     Optional[str] = None,
        client_id:    Optional[int] = None,
        deal_id:      Optional[int] = None,
        overdue_only: bool = False,
        limit:        int = 100,
        offset:       int = 0,
    ) -> List[Dict[str, Any]]:
        """
        Список задач.
        Поиск выполняется только по заголовку задачи, потому что именно он
        отображается в таблице UI и не должен давать «скрытых» совпадений по
        полю description.
        overdue_only=True — только просроченные (due_date < сегодня, status=pending).
        """
        conditions: List[str] = []
        params: List[Any] = []

        if status:
            conditions.append("status = ?")
            params.append(status)
        if priority:
            conditions.append("priority = ?")
            params.append(priority)
        if client_id is not None:
            conditions.append("client_id = ?")
            params.append(client_id)
        if deal_id is not None:
            conditions.append("deal_id = ?")
            params.append(deal_id)
        if overdue_only:
            conditions.append("due_date < ? AND status = 'pending'")
            params.append(_now()[:10])  # YYYY-MM-DD

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        with self._connect() as conn:
            rows = conn.execute(
                f"""
                SELECT * FROM tasks {where}
                ORDER BY
                    CASE priority WHEN 'high' THEN 0 WHEN 'medium' THEN 1 ELSE 2 END,
                    due_date ASC NULLS LAST
                """,
                params,
            ).fetchall()
            items = [dict(r) for r in rows]

        if search:
            needle = search.casefold()
            items = [
                item for item in items
                if _contains_text(item.get("title"), needle)
            ]

        return items[offset: offset + limit]

    def update_task(
        self, task_id: int, data: TaskUpdate
    ) -> Optional[Dict[str, Any]]:
        updates = data.model_dump(exclude_unset=True)
        if not updates:
            return self.get_task(task_id)

        updates["updated_at"] = _now()
        fields = ", ".join(f"{k} = ?" for k in updates)
        values = list(updates.values()) + [task_id]

        with self._connect() as conn:
            conn.execute(f"UPDATE tasks SET {fields} WHERE id = ?", values)
            row = conn.execute(
                "SELECT * FROM tasks WHERE id = ?", (task_id,)
            ).fetchone()
            return _row(row)

    def complete_task(self, task_id: int) -> Optional[Dict[str, Any]]:
        """Отметить задачу как выполненную."""
        with self._connect() as conn:
            conn.execute(
                "UPDATE tasks SET status = 'done', updated_at = ? WHERE id = ?",
                (_now(), task_id),
            )
            row = conn.execute(
                "SELECT * FROM tasks WHERE id = ?", (task_id,)
            ).fetchone()
            return _row(row)

    def delete_task(self, task_id: int) -> bool:
        with self._connect() as conn:
            cur = conn.execute("DELETE FROM tasks WHERE id = ?", (task_id,))
            return cur.rowcount > 0

    # ══════════════════════════════════════════════════════════════════════════
    #  DASHBOARD / STATS
    # ══════════════════════════════════════════════════════════════════════════

    def get_dashboard_stats(self) -> Dict[str, Any]:
        """Сводная статистика CRM для главного дашборда."""
        today = _now()[:10]
        with self._connect() as conn:
            c = conn.execute
            return {
                "clients": {
                    "total":    c("SELECT COUNT(*) FROM clients WHERE status != 'archived'").fetchone()[0],
                    "leads":    c("SELECT COUNT(*) FROM clients WHERE status = 'lead'").fetchone()[0],
                    "active":   c("SELECT COUNT(*) FROM clients WHERE status = 'active'").fetchone()[0],
                    "vip":      c("SELECT COUNT(*) FROM clients WHERE status = 'vip'").fetchone()[0],
                    "archived": c("SELECT COUNT(*) FROM clients WHERE status = 'archived'").fetchone()[0],
                },
                "deals": {
                    "total":         c("SELECT COUNT(*) FROM deals").fetchone()[0],
                    "open":          c("SELECT COUNT(*) FROM deals WHERE status IN ('new','in_progress')").fetchone()[0],
                    "won":           c("SELECT COUNT(*) FROM deals WHERE status = 'won'").fetchone()[0],
                    "lost":          c("SELECT COUNT(*) FROM deals WHERE status = 'lost'").fetchone()[0],
                    "total_revenue": c("SELECT COALESCE(SUM(amount),0) FROM deals WHERE status = 'won'").fetchone()[0],
                },
                "tasks": {
                    "pending":   c("SELECT COUNT(*) FROM tasks WHERE status = 'pending'").fetchone()[0],
                    "done":      c("SELECT COUNT(*) FROM tasks WHERE status = 'done'").fetchone()[0],
                    "overdue":   c(
                        "SELECT COUNT(*) FROM tasks WHERE status = 'pending' AND due_date < ?",
                        (today,),
                    ).fetchone()[0],
                    "high_prio": c(
                        "SELECT COUNT(*) FROM tasks WHERE status = 'pending' AND priority = 'high'"
                    ).fetchone()[0],
                },
            }
