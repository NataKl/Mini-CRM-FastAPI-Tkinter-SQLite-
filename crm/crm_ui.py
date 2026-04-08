"""
crm_ui.py — Tkinter-десктоп для мини-CRM
─────────────────────────────────────────
Делает HTTP-запросы к локальному FastAPI-серверу (crm_api.py).

Запуск:
  # 1. Запустите сервер (в отдельном терминале):
  venv\\Scripts\\python.exe crm\\crm.py

  # 2. Запустите UI:
  venv\\Scripts\\python.exe crm\\crm_ui.py
"""

import json
import sys
import threading
import webbrowser
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional

try:
    import requests
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "-q"])
    import requests

API_URL = "http://127.0.0.1:8000"

# ── Цветовая палитра ───────────────────────────────────────────────────────────
BG     = "#F0F2F5"
WHITE  = "#FFFFFF"
DARK   = "#2C3E50"
BLUE   = "#2980B9"
GREEN  = "#27AE60"
RED    = "#C0392B"
AMBER  = "#D4840A"
PURPLE = "#7D3C98"
GRAY   = "#6B7280"
STRIPE = "#F7F8FA"

CLIENT_STATUSES = ["lead", "active", "vip", "inactive", "archived"]
DEAL_STATUSES   = ["new", "in_progress", "won", "lost", "on_hold"]
TASK_STATUSES   = ["pending", "done", "cancelled"]
TASK_PRIORITIES = ["low", "medium", "high"]
CURRENCIES      = ["RUB", "USD", "EUR", "KZT", "GBP"]

STATUS_LABELS: Dict[str, str] = {
    "lead": "Лид", "active": "Активный", "vip": "VIP",
    "inactive": "Неактивный", "archived": "Архив",
    "new": "Новая", "in_progress": "В работе", "won": "Выиграна",
    "lost": "Проиграна", "on_hold": "На паузе",
    "pending": "Ожидает", "done": "Выполнено", "cancelled": "Отменено",
    "low": "Низкий", "medium": "Средний", "high": "Высокий",
}


# ── Файл настроек Google-интеграции ───────────────────────────────────────────
_GOOGLE_CFG = Path(__file__).parent / "google_settings.json"


def _load_google_settings() -> Dict[str, str]:
    try:
        if _GOOGLE_CFG.exists():
            return json.loads(_GOOGLE_CFG.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _save_google_settings(s: Dict[str, str]) -> None:
    cfg = _load_google_settings()
    cfg.update(s)
    _GOOGLE_CFG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


# ══════════════════════════════════════════════════════════════════════════════
#  Диалог настроек Google-интеграции
# ══════════════════════════════════════════════════════════════════════════════

class GoogleSettingsDialog(tk.Toplevel):
    """Диалог настроек Google-интеграции (сервисный аккаунт / OAuth2 / папка)."""

    _INSTRUCTIONS = (
        "Для работы экспорта отчётов необходимо:\n\n"
        "1. Создать сервисный аккаунт в Google Cloud Console\n"
        "2. Скачать JSON ключ сервисного аккаунта\n"
        "3. Создать OAuth2 клиент и скачать client_secret.json\n"
        "4. При первом использовании будет открыт браузер для авторизации\n"
        "5. (Опционально) Указать ID папки в Google Drive для сохранения отчётов\n"
        "6. Выбранные JSON-файлы и ID папки сохраняются в google_settings.json"
    )

    def __init__(self, parent: tk.Widget) -> None:
        super().__init__(parent)
        self.title("Настройки Google интеграции")
        self.resizable(False, False)
        self.configure(bg=WHITE)
        self.grab_set()
        self._build()
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.wait_window()

    def _build(self) -> None:
        cfg = _load_google_settings()
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill=tk.BOTH, expand=True)
        frm.columnconfigure(1, weight=1)

        # ── строка 0-1: сервисный аккаунт ────────────────────────────────────
        ttk.Label(frm, text="Путь к файлу сервисного аккаунта (JSON):").grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 2))
        self._v_sa = tk.StringVar(value=cfg.get("service_account_json", ""))
        ttk.Entry(frm, textvariable=self._v_sa, width=58).grid(
            row=1, column=0, sticky=tk.EW, padx=(0, 4))
        ttk.Button(frm, text="Обзор",
                   command=lambda: self._browse(self._v_sa)).grid(row=1, column=1, sticky=tk.W)

        # ── строка 2-3: client_secret ─────────────────────────────────────────
        ttk.Label(frm, text="Путь к client_secret файлу (JSON):").grid(
            row=2, column=0, columnspan=2, sticky=tk.W, pady=(8, 2))
        self._v_cs = tk.StringVar(value=cfg.get("client_secret_json", ""))
        ttk.Entry(frm, textvariable=self._v_cs, width=58).grid(
            row=3, column=0, sticky=tk.EW, padx=(0, 4))
        ttk.Button(frm, text="Обзор",
                   command=lambda: self._browse(self._v_cs)).grid(row=3, column=1, sticky=tk.W)

        # ── строка 4-5: folder ID ────────────────────────────────────────────
        ttk.Label(frm, text="ID папки в Google Drive (опционально):").grid(
            row=4, column=0, columnspan=2, sticky=tk.W, pady=(8, 2))
        self._v_folder = tk.StringVar(value=cfg.get("folder_id", ""))
        ttk.Entry(frm, textvariable=self._v_folder, width=50).grid(
            row=5, column=0, sticky=tk.EW, padx=(0, 4))
        ttk.Button(frm, text="Вставить",
                   command=lambda: self._paste(self._v_folder)).grid(row=5, column=1, sticky=tk.W)

        # ── строка 6: инструкция ─────────────────────────────────────────────
        inst = tk.Frame(frm, bg="#EEF1F5", bd=1, relief=tk.FLAT)
        inst.grid(row=6, column=0, columnspan=2, sticky=tk.EW, pady=(12, 8))
        tk.Label(inst, text=self._INSTRUCTIONS, bg="#EEF1F5", justify=tk.LEFT,
                 font=("TkDefaultFont", 8), wraplength=480,
                 anchor=tk.W).pack(padx=10, pady=8, anchor=tk.W)

        # ── строка 7: кнопки ─────────────────────────────────────────────────
        bf = ttk.Frame(frm)
        bf.grid(row=7, column=0, columnspan=2, pady=(4, 0))
        ttk.Button(bf, text="Сохранить", command=self._save, width=12).pack(side=tk.LEFT, padx=4)
        ttk.Button(bf, text="Отмена", command=self.destroy, width=10).pack(side=tk.LEFT)

    def _browse(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="Выберите JSON файл",
            filetypes=[("JSON файлы", "*.json"), ("Все файлы", "*.*")],
        )
        if path:
            var.set(path)

    def _paste(self, var: tk.StringVar) -> None:
        """Вставить текст из буфера обмена в поле var (кнопка-замена Ctrl+V)."""
        try:
            text = self.clipboard_get()
            var.set(text.strip())
        except Exception:
            pass

    def _save(self) -> None:
        _save_google_settings({
            "service_account_json": self._v_sa.get().strip(),
            "client_secret_json":   self._v_cs.get().strip(),
            "folder_id":            self._v_folder.get().strip(),
        })
        messagebox.showinfo("Настройки Google", "Настройки сохранены.", parent=self)
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  Экспорт отчёта в Google Sheets
# ══════════════════════════════════════════════════════════════════════════════

def _export_to_gsheet(parent: tk.Widget, sheet_title: str,
                      headers: List[str], rows: List[tuple],
                      summary: Optional[List[tuple]] = None) -> None:
    """Создаёт Google Sheet с данными и открывает ссылку в браузере.

    Структура листа:
      • Строка 1     — заголовок отчёта (широкий, тёмный фон)
      • Строки 2-N   — СВОДНАЯ ИНФОРМАЦИЯ (если передан summary)
      • Строка N+1   — ПОДРОБНЫЙ СПИСОК (заголовок секции)
      • Строка N+2   — шапка таблицы (синий фон)
      • Строки N+3+  — данные

    Запускается в отдельном потоке, чтобы не блокировать UI.
    При первом запуске открывает браузер для OAuth2-авторизации.

    Использует GUI-интеграцию из:
      - google_integration/google_drive.py
      - google_integration/google_sheets.py
    """
    cfg = _load_google_settings()

    client_secret = cfg.get("client_secret_json", "").strip()
    if not client_secret or not Path(client_secret).exists():
        messagebox.showerror(
            "Настройки Google",
            "Файл client_secret.json не найден.\n"
            "Откройте «Настройки Google» и укажите путь к файлу.",
            parent=parent,
        )
        return

    folder_id  = cfg.get("folder_id", "").strip() or None
    token_file = str(Path(__file__).parent.parent / "token.json")

    def _run() -> None:
        try:
            _project_root = str(Path(__file__).parent.parent)
            if _project_root not in sys.path:
                sys.path.insert(0, _project_root)

            from google_integration.google_drive import GoogleDriveIntegration  # noqa: PLC0415
            from google_integration.google_sheets import GoogleSheetsIntegration  # noqa: PLC0415

            drive_client = GoogleDriveIntegration(
                client_secret_file=client_secret,
                token_file=token_file,
                folder_id=folder_id,
            )
            file = drive_client.create_spreadsheet(sheet_title)
            sid  = file["id"]
            link = file.get("webViewLink",
                            f"https://docs.google.com/spreadsheets/d/{sid}")

            sheet_client = GoogleSheetsIntegration(
                spreadsheet_id=sid,
                client_secret_file=client_secret,
                token_file=token_file,
            )
            sheet_client.write_table_report(
                sheet_title=sheet_title,
                headers=list(headers),
                rows=[list(row) for row in rows],
                summary=list(summary or []),
            )

            _save_google_settings({
                "last_report_spreadsheet_id": sid,
                "last_report_title": sheet_title,
                "last_report_url": link,
            })

            parent.after(0, lambda: _show_export_success(parent, sheet_title, link))

        except Exception as exc:
            err = str(exc)
            parent.after(0, lambda: messagebox.showerror(
                "Ошибка экспорта", err, parent=parent))

    threading.Thread(target=_run, daemon=True).start()


def _show_export_success(parent: tk.Widget, title: str, link: str) -> None:
    """Диалог с результатом — ссылкой на созданный Google Sheet."""
    dlg = tk.Toplevel(parent)
    dlg.title("Отчёт создан")
    dlg.resizable(False, False)
    dlg.configure(bg=WHITE)
    dlg.grab_set()

    ttk.Label(dlg, text="Google Sheet успешно создан:",
              padding=(16, 14, 16, 4)).pack()
    ttk.Label(dlg, text=title,
              font=("TkDefaultFont", 10, "bold"),
              padding=(16, 0, 16, 6)).pack()

    ttk.Label(dlg, text="Ссылка на отчёт:",
              padding=(16, 4, 16, 2)).pack(anchor=tk.W)
    lnk_lbl = tk.Label(dlg, text=link, fg=BLUE, bg=WHITE, cursor="hand2",
                        font=("TkDefaultFont", 8), wraplength=400)
    lnk_lbl.pack(padx=16, pady=(0, 4))
    lnk_lbl.bind("<Button-1>", lambda _: webbrowser.open(link))

    bf = ttk.Frame(dlg, padding=(16, 8, 16, 14))
    bf.pack()
    ttk.Button(
        bf,
        text="Скопировать ссылку",
        command=lambda: (dlg.clipboard_clear(), dlg.clipboard_append(link)),
        width=20,
    ).pack(side=tk.LEFT, padx=4)
    ttk.Button(bf, text="Открыть в браузере",
               command=lambda: webbrowser.open(link), width=20).pack(side=tk.LEFT, padx=4)
    ttk.Button(bf, text="Закрыть", command=dlg.destroy, width=10).pack(side=tk.LEFT)

    dlg.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width()  - dlg.winfo_width())  // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - dlg.winfo_height()) // 2
    dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")


# ══════════════════════════════════════════════════════════════════════════════
#  HTTP-клиент
# ══════════════════════════════════════════════════════════════════════════════

class APIError(Exception):
    pass


class API:
    """Тонкая обёртка над requests для CRM REST API."""

    def __init__(self, base_url: str = API_URL) -> None:
        self.base = base_url.rstrip("/")
        self._s = requests.Session()
        self._s.headers["Content-Type"] = "application/json"

    def _url(self, path: str) -> str:
        return f"{self.base}/{path.lstrip('/')}"

    def _handle(self, resp: "requests.Response") -> Any:
        if resp.status_code == 204:
            return None
        if resp.ok:
            return resp.json()
        try:
            detail = resp.json().get("detail", resp.text)
        except Exception:
            detail = resp.text
        raise APIError(f"HTTP {resp.status_code}: {detail}")

    def get(self, path: str, **params) -> Any:
        p = {k: v for k, v in params.items() if v is not None and v != ""}
        return self._handle(self._s.get(self._url(path), params=p or None, timeout=5))

    def post(self, path: str, body: Dict) -> Any:
        return self._handle(self._s.post(self._url(path), data=json.dumps(body), timeout=5))

    def put(self, path: str, body: Dict) -> Any:
        return self._handle(self._s.put(self._url(path), data=json.dumps(body), timeout=5))

    def patch(self, path: str, **params) -> Any:
        p = {k: v for k, v in params.items() if v is not None and v != ""}
        return self._handle(self._s.patch(self._url(path), params=p or None, timeout=5))

    def delete(self, path: str) -> None:
        return self._handle(self._s.delete(self._url(path), timeout=5))

    def ping(self) -> bool:
        try:
            self._s.get(self._url("/dashboard"), timeout=2)
            return True
        except Exception:
            return False


# ══════════════════════════════════════════════════════════════════════════════
#  Вспомогательные функции
# ══════════════════════════════════════════════════════════════════════════════

def _fmt_dt(iso: Optional[str]) -> str:
    """'2026-04-07T19:33:45' → '07.04.2026 19:33'"""
    if not iso:
        return ""
    try:
        dt = iso[:16]
        date, time = dt.split("T")
        y, m, d = date.split("-")
        return f"{d}.{m}.{y} {time}"
    except Exception:
        return iso[:16] if iso else ""


def _fmt_money(amount: Any, currency: str) -> str:
    try:
        v = float(amount)
        if v == 0:
            return "—"
        return f"{v:,.0f} {currency}".replace(",", " ")
    except Exception:
        return str(amount)


def _lbl(status: str) -> str:
    return STATUS_LABELS.get(status, status)


# ══════════════════════════════════════════════════════════════════════════════
#  Базовый модальный диалог (форма)
# ══════════════════════════════════════════════════════════════════════════════

class _FormDialog(tk.Toplevel):
    """
    Базовый класс для модальных форм создания / редактирования.

    Подкласс должен:
      - До вызова super().__init__() инициализировать нужные StringVar / BooleanVar.
      - Переопределить _build() для размещения полей формы.
      - Переопределить _collect() → Dict | None (None = валидация не прошла).
      - Переопределить _api_call(data) для вызова нужного метода API.
    """

    def __init__(self, parent: tk.Widget, title: str, api: API) -> None:
        super().__init__(parent)
        self.api = api
        self.result: Optional[Dict] = None

        self.title(title)
        self.resizable(False, False)
        self.configure(bg=WHITE)
        self.grab_set()

        self._body = ttk.Frame(self, padding=16)
        self._body.pack(fill=tk.BOTH, expand=True)
        self._body.columnconfigure(1, weight=1)

        self._build()
        self._add_buttons()
        self.update_idletasks()
        self._center(parent)
        self.wait_window()

    # ── layout helpers ────────────────────────────────────────────────────────

    def _entry(self, row: int, label: str, var: tk.StringVar, width: int = 34) -> ttk.Entry:
        ttk.Label(self._body, text=label).grid(row=row, column=0, sticky=tk.W, pady=4, padx=(0, 6))
        e = ttk.Entry(self._body, textvariable=var, width=width)
        e.grid(row=row, column=1, sticky=tk.EW, pady=4)
        return e

    def _combo(self, row: int, label: str, var: tk.StringVar,
               values: List[str], width: int = 32) -> ttk.Combobox:
        ttk.Label(self._body, text=label).grid(row=row, column=0, sticky=tk.W, pady=4, padx=(0, 6))
        c = ttk.Combobox(self._body, textvariable=var, values=values,
                         state="readonly", width=width)
        c.grid(row=row, column=1, sticky=tk.EW, pady=4)
        return c

    def _textarea(self, row: int, label: str, height: int = 3) -> tk.Text:
        ttk.Label(self._body, text=label).grid(row=row, column=0, sticky=tk.NW, pady=4, padx=(0, 6))
        t = tk.Text(self._body, width=34, height=height, wrap=tk.WORD,
                    relief=tk.SOLID, borderwidth=1, font=("TkDefaultFont", 9))
        t.grid(row=row, column=1, sticky=tk.EW, pady=4)
        return t

    def _add_buttons(self) -> None:
        frm = ttk.Frame(self, padding=(16, 4, 16, 12))
        frm.pack(fill=tk.X)
        ttk.Button(frm, text="Отмена",    command=self.destroy,   width=10).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(frm, text="Сохранить", command=self._on_save, width=10).pack(side=tk.RIGHT)

    def _center(self, parent: tk.Widget) -> None:
        x = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    # ── overridable ───────────────────────────────────────────────────────────

    def _build(self) -> None:
        pass

    def _collect(self) -> Optional[Dict]:
        return None

    def _api_call(self, data: Dict) -> Dict:
        raise NotImplementedError

    def _on_save(self) -> None:
        data = self._collect()
        if data is None:
            return
        try:
            self.result = self._api_call(data)
            self.destroy()
        except APIError as e:
            messagebox.showerror("Ошибка API", str(e), parent=self)
        except requests.exceptions.ConnectionError:
            messagebox.showerror("Нет соединения", "Сервер недоступен.\nЗапустите crm.py.", parent=self)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e), parent=self)


# ══════════════════════════════════════════════════════════════════════════════
#  Диалог: Клиент
# ══════════════════════════════════════════════════════════════════════════════

class ClientDialog(_FormDialog):

    def __init__(self, parent: tk.Widget, api: API,
                 existing: Optional[Dict] = None) -> None:
        self._ex = existing or {}
        self._v_name    = tk.StringVar(value=self._ex.get("name", ""))
        self._v_email   = tk.StringVar(value=self._ex.get("email", "") or "")
        self._v_phone   = tk.StringVar(value=self._ex.get("phone", "") or "")
        self._v_company = tk.StringVar(value=self._ex.get("company", "") or "")
        self._v_source  = tk.StringVar(value=self._ex.get("source", "") or "")
        self._v_status  = tk.StringVar(value=self._ex.get("status", "lead"))
        super().__init__(parent, "Редактировать клиента" if existing else "Новый клиент", api)

    def _build(self) -> None:
        self._entry(0, "Имя *",      self._v_name)
        self._entry(1, "Email",      self._v_email)
        self._entry(2, "Телефон",    self._v_phone)
        self._entry(3, "Компания",   self._v_company)
        self._entry(4, "Источник",   self._v_source)
        self._combo(5, "Статус",     self._v_status, CLIENT_STATUSES)
        self._notes_w = self._textarea(6, "Заметки")
        if self._ex.get("notes"):
            self._notes_w.insert("1.0", self._ex["notes"])

    def _collect(self) -> Optional[Dict]:
        name = self._v_name.get().strip()
        if not name:
            messagebox.showwarning("Обязательное поле", "Поле «Имя» не может быть пустым.", parent=self)
            return None
        return {
            "name":    name,
            "email":   self._v_email.get().strip()   or None,
            "phone":   self._v_phone.get().strip()   or None,
            "company": self._v_company.get().strip() or None,
            "source":  self._v_source.get().strip()  or None,
            "status":  self._v_status.get(),
            "notes":   self._notes_w.get("1.0", tk.END).strip() or None,
        }

    def _api_call(self, data: Dict) -> Dict:
        return (self.api.put(f"/clients/{self._ex['id']}", data)
                if self._ex else self.api.post("/clients", data))


# ══════════════════════════════════════════════════════════════════════════════
#  Диалог: Сделка
# ══════════════════════════════════════════════════════════════════════════════

class DealDialog(_FormDialog):

    def __init__(self, parent: tk.Widget, api: API,
                 existing: Optional[Dict] = None) -> None:
        self._ex = existing or {}
        # загрузить клиентов для выпадающего списка
        try:
            clients = api.get("/clients", limit=500)
        except Exception:
            clients = []
        self._client_map: Dict[str, Optional[int]] = {"— не указан —": None}
        self._client_opts = ["— не указан —"]
        for c in clients:
            lbl = f"{c['name']} (#{c['id']})"
            self._client_map[lbl] = c["id"]
            self._client_opts.append(lbl)

        ex_cid = self._ex.get("client_id")
        ex_c_lbl = next((l for l, v in self._client_map.items() if v == ex_cid),
                        "— не указан —")

        self._v_title    = tk.StringVar(value=self._ex.get("title", ""))
        self._v_client   = tk.StringVar(value=ex_c_lbl)
        self._v_amount   = tk.StringVar(value=str(self._ex.get("amount", 0)))
        self._v_currency = tk.StringVar(value=self._ex.get("currency", "RUB"))
        self._v_status   = tk.StringVar(value=self._ex.get("status", "new"))
        self._v_due      = tk.StringVar(value=self._ex.get("due_date", "") or "")
        super().__init__(parent, "Редактировать сделку" if existing else "Новая сделка", api)

    def _build(self) -> None:
        self._entry(0, "Название *",       self._v_title)
        self._combo(1, "Клиент",           self._v_client,   self._client_opts)
        self._entry(2, "Сумма",            self._v_amount)
        self._combo(3, "Валюта",           self._v_currency, CURRENCIES)
        self._combo(4, "Статус",           self._v_status,   DEAL_STATUSES)
        self._entry(5, "Срок (ГГГГ-ММ-ДД)", self._v_due)
        self._desc_w = self._textarea(6, "Описание")
        if self._ex.get("description"):
            self._desc_w.insert("1.0", self._ex["description"])

    def _collect(self) -> Optional[Dict]:
        title = self._v_title.get().strip()
        if not title:
            messagebox.showwarning("Обязательное поле", "Поле «Название» не может быть пустым.", parent=self)
            return None
        try:
            amount = float(self._v_amount.get() or 0)
        except ValueError:
            messagebox.showwarning("Валидация", "Сумма должна быть числом.", parent=self)
            return None
        return {
            "title":       title,
            "client_id":   self._client_map.get(self._v_client.get()),
            "amount":      amount,
            "currency":    self._v_currency.get(),
            "status":      self._v_status.get(),
            "due_date":    self._v_due.get().strip() or None,
            "description": self._desc_w.get("1.0", tk.END).strip() or None,
        }

    def _api_call(self, data: Dict) -> Dict:
        return (self.api.put(f"/deals/{self._ex['id']}", data)
                if self._ex else self.api.post("/deals", data))


# ══════════════════════════════════════════════════════════════════════════════
#  Диалог: Задача
# ══════════════════════════════════════════════════════════════════════════════

class TaskDialog(_FormDialog):

    def __init__(self, parent: tk.Widget, api: API,
                 existing: Optional[Dict] = None) -> None:
        self._ex = existing or {}
        try:
            clients = api.get("/clients", limit=500)
        except Exception:
            clients = []
        try:
            deals = api.get("/deals", limit=500)
        except Exception:
            deals = []

        self._client_map: Dict[str, Optional[int]] = {"— не указан —": None}
        self._client_opts = ["— не указан —"]
        for c in clients:
            lbl = f"{c['name']} (#{c['id']})"
            self._client_map[lbl] = c["id"]
            self._client_opts.append(lbl)

        self._deal_map: Dict[str, Optional[int]] = {"— не указана —": None}
        self._deal_opts = ["— не указана —"]
        for d in deals:
            lbl = f"{d['title']} (#{d['id']})"
            self._deal_map[lbl] = d["id"]
            self._deal_opts.append(lbl)

        ex_cid = self._ex.get("client_id")
        ex_did = self._ex.get("deal_id")
        ex_c = next((l for l, v in self._client_map.items() if v == ex_cid), "— не указан —")
        ex_d = next((l for l, v in self._deal_map.items()   if v == ex_did), "— не указана —")

        self._v_title    = tk.StringVar(value=self._ex.get("title", ""))
        self._v_client   = tk.StringVar(value=ex_c)
        self._v_deal     = tk.StringVar(value=ex_d)
        self._v_priority = tk.StringVar(value=self._ex.get("priority", "medium"))
        self._v_status   = tk.StringVar(value=self._ex.get("status", "pending"))
        self._v_due      = tk.StringVar(value=self._ex.get("due_date", "") or "")
        super().__init__(parent, "Редактировать задачу" if existing else "Новая задача", api)

    def _build(self) -> None:
        self._entry(0, "Задача *",           self._v_title)
        self._combo(1, "Клиент",             self._v_client,   self._client_opts)
        self._combo(2, "Сделка",             self._v_deal,     self._deal_opts)
        self._combo(3, "Приоритет",          self._v_priority, TASK_PRIORITIES)
        self._combo(4, "Статус",             self._v_status,   TASK_STATUSES)
        self._entry(5, "Срок (ГГГГ-ММ-ДД)", self._v_due)
        self._desc_w = self._textarea(6, "Описание")
        if self._ex.get("description"):
            self._desc_w.insert("1.0", self._ex["description"])

    def _collect(self) -> Optional[Dict]:
        title = self._v_title.get().strip()
        if not title:
            messagebox.showwarning("Обязательное поле", "Поле «Задача» не может быть пустым.", parent=self)
            return None
        return {
            "title":       title,
            "client_id":   self._client_map.get(self._v_client.get()),
            "deal_id":     self._deal_map.get(self._v_deal.get()),
            "priority":    self._v_priority.get(),
            "status":      self._v_status.get(),
            "due_date":    self._v_due.get().strip() or None,
            "description": self._desc_w.get("1.0", tk.END).strip() or None,
        }

    def _api_call(self, data: Dict) -> Dict:
        return (self.api.put(f"/tasks/{self._ex['id']}", data)
                if self._ex else self.api.post("/tasks", data))


# ══════════════════════════════════════════════════════════════════════════════
#  Вспомогательный диалог смены статуса (Deals / Tasks)
# ══════════════════════════════════════════════════════════════════════════════

def _ask_status(parent: tk.Widget, title: str, options: List[str],
                current: str = "") -> Optional[str]:
    """Маленький диалог выбора статуса из списка. Возвращает выбранное значение или None."""
    result: List[Optional[str]] = [None]
    dlg = tk.Toplevel(parent)
    dlg.title(title)
    dlg.resizable(False, False)
    dlg.grab_set()
    var = tk.StringVar(value=current or options[0])
    ttk.Label(dlg, text="Выберите статус:", padding=(12, 10, 12, 4)).pack()
    cb = ttk.Combobox(dlg, textvariable=var, values=options, state="readonly", width=22)
    cb.pack(padx=16, pady=4)
    frm = ttk.Frame(dlg, padding=(12, 6, 12, 12))
    frm.pack()

    def ok():
        result[0] = var.get()
        dlg.destroy()

    ttk.Button(frm, text="Применить", command=ok,          width=10).pack(side=tk.LEFT, padx=4)
    ttk.Button(frm, text="Отмена",   command=dlg.destroy, width=10).pack(side=tk.LEFT)
    dlg.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width()  - dlg.winfo_width())  // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - dlg.winfo_height()) // 2
    dlg.geometry(f"+{max(x,0)}+{max(y,0)}")
    dlg.wait_window()
    return result[0]


# ══════════════════════════════════════════════════════════════════════════════
#  Базовый фрейм вкладки с Treeview
# ══════════════════════════════════════════════════════════════════════════════

class _TableFrame(ttk.Frame):
    """Базовый класс для вкладок с таблицей (Treeview) и CRUD-кнопками."""

    COLS:   List[str] = []
    HDRS:   List[str] = []
    WIDTHS: List[int] = []

    def __init__(self, parent: tk.Widget, api: API) -> None:
        super().__init__(parent, padding=10)
        self.api = api
        self._build_top()
        self._build_tree()
        self._build_bottom()
        self.refresh()

    # ── заглушки для переопределения ─────────────────────────────────────────

    def _build_top(self) -> None:
        """Панель фильтров/поиска."""

    def _build_bottom(self) -> None:
        """Панель кнопок."""

    def refresh(self) -> None:
        """Загрузить данные из API и обновить таблицу."""

    # ── treeview ──────────────────────────────────────────────────────────────

    def _build_tree(self) -> None:
        frm = ttk.Frame(self)
        frm.pack(fill=tk.BOTH, expand=True, pady=(6, 0))

        self.tv = ttk.Treeview(frm, columns=self.COLS, show="headings",
                                selectmode="browse")
        for col, hdr, w in zip(self.COLS, self.HDRS, self.WIDTHS):
            anchor = tk.CENTER if col == "id" else tk.W
            self.tv.heading(col, text=hdr)
            self.tv.column(col, width=w, minwidth=30, anchor=anchor)

        sb_y = ttk.Scrollbar(frm, orient=tk.VERTICAL,   command=self.tv.yview)
        sb_x = ttk.Scrollbar(frm, orient=tk.HORIZONTAL, command=self.tv.xview)
        self.tv.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        self.tv.grid(row=0, column=0, sticky=tk.NSEW)
        sb_y.grid(row=0, column=1, sticky=tk.NS)
        sb_x.grid(row=1, column=0, sticky=tk.EW)
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        self.tv.tag_configure("odd", background=STRIPE)
        self.tv.bind("<Double-1>", lambda _: self._on_dbl_click())

    def _on_dbl_click(self) -> None:
        pass

    # ── helpers ───────────────────────────────────────────────────────────────

    def _selected_id(self) -> Optional[int]:
        sel = self.tv.selection()
        return int(sel[0]) if sel else None

    def _fill_tree(self, rows: List[Dict], mapper) -> None:
        self.tv.delete(*self.tv.get_children())
        for i, row in enumerate(rows):
            tag = "odd" if i % 2 else ""
            self.tv.insert("", tk.END, iid=str(row["id"]),
                           tags=(tag,), values=mapper(row))

    def _count_label(self, lbl: tk.Label, n: int) -> None:
        lbl.config(text=f"Записей: {n}")

    def _err(self, exc: Exception) -> None:
        messagebox.showerror("Ошибка", str(exc), parent=self)

    def _confirm(self, msg: str) -> bool:
        return messagebox.askyesno("Подтверждение", msg, parent=self)


# ══════════════════════════════════════════════════════════════════════════════
#  Вкладка: Дашборд
# ══════════════════════════════════════════════════════════════════════════════

class DashboardFrame(ttk.Frame):

    def __init__(self, parent: tk.Widget, api: API) -> None:
        super().__init__(parent, padding=20)
        self.api = api
        self._vals: Dict[str, tk.Label] = {}
        self._build()
        self.refresh()

    def _build(self) -> None:
        ttk.Label(self, text="Дашборд",
                  font=("TkDefaultFont", 14, "bold")).pack(anchor=tk.W)
        ttk.Label(self, text="Сводная статистика CRM",
                  foreground=GRAY).pack(anchor=tk.W, pady=(0, 16))

        cards = ttk.Frame(self)
        cards.pack(fill=tk.X)

        sections = [
            ("Клиенты", [
                ("total",    "Всего (без архива)"),
                ("leads",    "Лиды"),
                ("active",   "Активные"),
                ("vip",      "VIP"),
                ("archived", "В архиве"),
            ]),
            ("Сделки", [
                ("total",         "Всего"),
                ("open",          "Открытых"),
                ("won",           "Выиграно"),
                ("lost",          "Проиграно"),
                ("total_revenue", "Выручка (won)"),
            ]),
            ("Задачи", [
                ("pending",   "Ожидают"),
                ("done",      "Выполнено"),
                ("overdue",   "Просрочено"),
                ("high_prio", "Высокий приоритет"),
            ]),
        ]

        for col, (section, items) in enumerate(sections):
            card = ttk.LabelFrame(cards, text=section, padding=12)
            card.grid(row=0, column=col, sticky=tk.NSEW, padx=6, pady=4)
            cards.columnconfigure(col, weight=1)
            for key, label in items:
                row_f = ttk.Frame(card)
                row_f.pack(fill=tk.X, pady=3)
                ttk.Label(row_f, text=label + ":", foreground=GRAY).pack(side=tk.LEFT)
                lbl = tk.Label(row_f, text="—", font=("TkDefaultFont", 10, "bold"),
                               bg=BG)
                lbl.pack(side=tk.RIGHT)
                self._vals[f"{section}.{key}"] = lbl

        ttk.Button(self, text="Обновить", command=self.refresh).pack(
            anchor=tk.W, pady=(18, 0))

    def refresh(self) -> None:
        try:
            s = self.api.get("/dashboard")
        except Exception:
            return
        mapping = {
            "Клиенты.total":    s["clients"]["total"],
            "Клиенты.leads":    s["clients"]["leads"],
            "Клиенты.active":   s["clients"]["active"],
            "Клиенты.vip":      s["clients"]["vip"],
            "Клиенты.archived": s["clients"]["archived"],
            "Сделки.total":         s["deals"]["total"],
            "Сделки.open":          s["deals"]["open"],
            "Сделки.won":           s["deals"]["won"],
            "Сделки.lost":          s["deals"]["lost"],
            "Сделки.total_revenue": f"{s['deals']['total_revenue']:,.0f}".replace(",", " "),
            "Задачи.pending":   s["tasks"]["pending"],
            "Задачи.done":      s["tasks"]["done"],
            "Задачи.overdue":   s["tasks"]["overdue"],
            "Задачи.high_prio": s["tasks"]["high_prio"],
        }
        for k, v in mapping.items():
            if k in self._vals:
                self._vals[k].config(text=str(v))


# ══════════════════════════════════════════════════════════════════════════════
#  Вкладка: Клиенты
# ══════════════════════════════════════════════════════════════════════════════

class ClientsFrame(_TableFrame):

    COLS   = ("id", "name", "email", "phone", "company", "status", "updated_at")
    HDRS   = ("ID",  "Имя",   "Email",  "Телефон", "Компания",  "Статус",   "Обновлён")
    WIDTHS = (40,    170,     165,      115,        145,          100,         130)

    def _build_top(self) -> None:
        top = ttk.Frame(self)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Поиск:").pack(side=tk.LEFT)
        self._v_search = tk.StringVar()
        e = ttk.Entry(top, textvariable=self._v_search, width=24)
        e.pack(side=tk.LEFT, padx=(4, 10))
        e.bind("<Return>", lambda _: self.refresh())

        ttk.Label(top, text="Статус:").pack(side=tk.LEFT)
        self._v_status = tk.StringVar()
        cb = ttk.Combobox(top, textvariable=self._v_status,
                          values=[""] + CLIENT_STATUSES, state="readonly", width=12)
        cb.pack(side=tk.LEFT, padx=(4, 10))
        cb.bind("<<ComboboxSelected>>", lambda _: self.refresh())

        ttk.Button(top, text="Найти",    command=self.refresh).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="Сбросить", command=self._reset).pack(side=tk.LEFT, padx=2)

    def _build_bottom(self) -> None:
        bot = ttk.Frame(self)
        bot.pack(fill=tk.X, pady=(6, 0))

        ttk.Button(bot, text="+ Добавить",    command=self._create).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Редактировать", command=self._edit).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Архивировать",  command=self._archive).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Удалить",       command=self._delete).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="↑ Выгрузить отчёт",
                   command=self._export_report).pack(side=tk.LEFT, padx=(10, 2))

        self._cnt = ttk.Label(bot, text="", foreground=GRAY)
        self._cnt.pack(side=tk.RIGHT)

    def _on_dbl_click(self) -> None:
        self._edit()

    def _reset(self) -> None:
        self._v_search.set("")
        self._v_status.set("")
        self.refresh()

    def refresh(self) -> None:
        try:
            data = self.api.get("/clients",
                                search=self._v_search.get().strip() or None,
                                status=self._v_status.get() or None,
                                limit=500)
        except Exception as e:
            self._err(e)
            return
        self._fill_tree(data, lambda r: (
            r["id"],
            r.get("name", ""),
            r.get("email", "") or "",
            r.get("phone", "") or "",
            r.get("company", "") or "",
            _lbl(r.get("status", "")),
            _fmt_dt(r.get("updated_at", "")),
        ))
        self._count_label(self._cnt, len(data))

    # ── CRUD ──────────────────────────────────────────────────────────────────

    def _create(self) -> None:
        if ClientDialog(self, self.api).result is not None:
            self.refresh()

    def _edit(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите клиента в списке.", parent=self)
            return
        try:
            item = self.api.get(f"/clients/{iid}")
        except APIError as e:
            self._err(e)
            return
        if ClientDialog(self, self.api, existing=item).result is not None:
            self.refresh()

    def _archive(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите клиента.", parent=self)
            return
        if self._confirm("Перевести клиента в архив?"):
            try:
                self.api.patch(f"/clients/{iid}/archive")
                self.refresh()
            except APIError as e:
                self._err(e)

    def _delete(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите клиента.", parent=self)
            return
        if self._confirm("Удалить клиента безвозвратно?"):
            try:
                self.api.delete(f"/clients/{iid}")
                self.refresh()
            except APIError as e:
                self._err(e)

    def _export_report(self) -> None:
        rows = [self.tv.item(iid)["values"] for iid in self.tv.get_children()]
        if not rows:
            messagebox.showinfo("Экспорт", "Нет данных для экспорта.", parent=self)
            return
        today = datetime.now().strftime("%d.%m.%Y")
        # Сводная информация из дашборда (раздел «Клиенты»)
        summary: List[tuple] = []
        try:
            s  = self.api.get("/dashboard")
            sc = s.get("clients", {})
            summary = [
                ("Всего клиентов (без архива)", sc.get("total", 0)),
                ("Лиды",                        sc.get("leads", 0)),
                ("Активные",                    sc.get("active", 0)),
                ("VIP",                         sc.get("vip", 0)),
                ("В архиве",                    sc.get("archived", 0)),
            ]
        except Exception:
            pass
        _export_to_gsheet(self, f"CRM — Клиенты {today}", list(self.HDRS), rows,
                          summary=summary)


# ══════════════════════════════════════════════════════════════════════════════
#  Вкладка: Сделки
# ══════════════════════════════════════════════════════════════════════════════

class DealsFrame(_TableFrame):

    COLS   = ("id", "title",    "client",  "amount",  "status",   "due_date", "updated_at")
    HDRS   = ("ID",  "Название", "Клиент",  "Сумма",   "Статус",   "Срок",     "Обновлён")
    WIDTHS = (40,    210,        155,        110,        105,         100,         130)

    def __init__(self, parent: tk.Widget, api: API) -> None:
        self._clients_cache: Dict[int, str] = {}
        super().__init__(parent, api)

    def _build_top(self) -> None:
        top = ttk.Frame(self)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Поиск:").pack(side=tk.LEFT)
        self._v_search = tk.StringVar()
        e = ttk.Entry(top, textvariable=self._v_search, width=24)
        e.pack(side=tk.LEFT, padx=(4, 10))
        e.bind("<Return>", lambda _: self.refresh())

        ttk.Label(top, text="Статус:").pack(side=tk.LEFT)
        self._v_status = tk.StringVar()
        cb = ttk.Combobox(top, textvariable=self._v_status,
                          values=[""] + DEAL_STATUSES, state="readonly", width=14)
        cb.pack(side=tk.LEFT, padx=(4, 10))
        cb.bind("<<ComboboxSelected>>", lambda _: self.refresh())

        ttk.Button(top, text="Найти",    command=self.refresh).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="Сбросить", command=self._reset).pack(side=tk.LEFT, padx=2)

    def _build_bottom(self) -> None:
        bot = ttk.Frame(self)
        bot.pack(fill=tk.X, pady=(6, 0))

        ttk.Button(bot, text="+ Добавить",    command=self._create).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Редактировать", command=self._edit).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Изменить статус", command=self._change_status).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Удалить",       command=self._delete).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="↑ Выгрузить отчёт",
                   command=self._export_report).pack(side=tk.LEFT, padx=(10, 2))

        self._cnt = ttk.Label(bot, text="", foreground=GRAY)
        self._cnt.pack(side=tk.RIGHT)

    def _on_dbl_click(self) -> None:
        self._edit()

    def _reset(self) -> None:
        self._v_search.set("")
        self._v_status.set("")
        self.refresh()

    def refresh(self) -> None:
        try:
            clients = self.api.get("/clients", limit=500)
            self._clients_cache = {c["id"]: c["name"] for c in clients}
        except Exception:
            pass
        try:
            data = self.api.get("/deals",
                                search=self._v_search.get().strip() or None,
                                status=self._v_status.get() or None,
                                limit=500)
        except Exception as e:
            self._err(e)
            return
        self._fill_tree(data, lambda r: (
            r["id"],
            r.get("title", ""),
            self._clients_cache.get(r.get("client_id"), "—"),
            _fmt_money(r.get("amount", 0), r.get("currency", "RUB")),
            _lbl(r.get("status", "")),
            r.get("due_date", "") or "—",
            _fmt_dt(r.get("updated_at", "")),
        ))
        self._count_label(self._cnt, len(data))

    # ── CRUD ──────────────────────────────────────────────────────────────────

    def _create(self) -> None:
        if DealDialog(self, self.api).result is not None:
            self.refresh()

    def _edit(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите сделку в списке.", parent=self)
            return
        try:
            item = self.api.get(f"/deals/{iid}")
        except APIError as e:
            self._err(e)
            return
        if DealDialog(self, self.api, existing=item).result is not None:
            self.refresh()

    def _change_status(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите сделку.", parent=self)
            return
        try:
            item = self.api.get(f"/deals/{iid}")
        except APIError as e:
            self._err(e)
            return
        new_status = _ask_status(self, "Статус сделки", DEAL_STATUSES,
                                  current=item.get("status", ""))
        if new_status:
            try:
                self.api.patch(f"/deals/{iid}/status", status=new_status)
                self.refresh()
            except APIError as e:
                self._err(e)

    def _delete(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите сделку.", parent=self)
            return
        if self._confirm("Удалить сделку безвозвратно?"):
            try:
                self.api.delete(f"/deals/{iid}")
                self.refresh()
            except APIError as e:
                self._err(e)

    def _export_report(self) -> None:
        rows = [self.tv.item(iid)["values"] for iid in self.tv.get_children()]
        if not rows:
            messagebox.showinfo("Экспорт", "Нет данных для экспорта.", parent=self)
            return
        today = datetime.now().strftime("%d.%m.%Y")
        # Сводная информация из дашборда (раздел «Сделки»)
        summary: List[tuple] = []
        try:
            s  = self.api.get("/dashboard")
            sd = s.get("deals", {})
            rev = sd.get("total_revenue", 0)
            summary = [
                ("Всего сделок",          sd.get("total", 0)),
                ("Открытых",              sd.get("open",  0)),
                ("Выиграно",              sd.get("won",   0)),
                ("Проиграно",             sd.get("lost",  0)),
                ("Общая выручка (won)",   f"{rev:,.0f}".replace(",", " ")),
            ]
        except Exception:
            pass
        _export_to_gsheet(self, f"CRM — Сделки {today}", list(self.HDRS), rows,
                          summary=summary)


# ══════════════════════════════════════════════════════════════════════════════
#  Вкладка: Задачи
# ══════════════════════════════════════════════════════════════════════════════

class TasksFrame(_TableFrame):

    COLS   = ("id", "title",   "priority",   "status",   "due_date", "client",  "deal",    "updated_at")
    HDRS   = ("ID",  "Задача",  "Приоритет",  "Статус",   "Срок",     "Клиент",  "Сделка",  "Обновлён")
    WIDTHS = (40,    205,        90,            90,          100,         130,       130,        130)

    def __init__(self, parent: tk.Widget, api: API) -> None:
        self._clients_cache: Dict[int, str] = {}
        self._deals_cache:   Dict[int, str] = {}
        super().__init__(parent, api)

    def _build_top(self) -> None:
        top = ttk.Frame(self)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Поиск:").pack(side=tk.LEFT)
        self._v_search = tk.StringVar()
        e = ttk.Entry(top, textvariable=self._v_search, width=20)
        e.pack(side=tk.LEFT, padx=(4, 8))
        e.bind("<Return>", lambda _: self.refresh())

        ttk.Label(top, text="Статус:").pack(side=tk.LEFT)
        self._v_status = tk.StringVar()
        cb_s = ttk.Combobox(top, textvariable=self._v_status,
                            values=[""] + TASK_STATUSES, state="readonly", width=11)
        cb_s.pack(side=tk.LEFT, padx=(4, 8))
        cb_s.bind("<<ComboboxSelected>>", lambda _: self.refresh())

        ttk.Label(top, text="Приоритет:").pack(side=tk.LEFT)
        self._v_prio = tk.StringVar()
        cb_p = ttk.Combobox(top, textvariable=self._v_prio,
                            values=[""] + TASK_PRIORITIES, state="readonly", width=9)
        cb_p.pack(side=tk.LEFT, padx=(4, 8))
        cb_p.bind("<<ComboboxSelected>>", lambda _: self.refresh())

        self._v_overdue = tk.BooleanVar()
        ttk.Checkbutton(top, text="Просрочен.", variable=self._v_overdue,
                        command=self.refresh).pack(side=tk.LEFT, padx=4)

        ttk.Button(top, text="Найти",    command=self.refresh).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="Сбросить", command=self._reset).pack(side=tk.LEFT, padx=2)

    def _build_bottom(self) -> None:
        bot = ttk.Frame(self)
        bot.pack(fill=tk.X, pady=(6, 0))

        ttk.Button(bot, text="+ Добавить",    command=self._create).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Редактировать", command=self._edit).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Выполнено",     command=self._complete).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Изменить статус", command=self._change_status).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="Удалить",       command=self._delete).pack(side=tk.LEFT, padx=2)
        ttk.Button(bot, text="↑ Выгрузить отчёт",
                   command=self._export_report).pack(side=tk.LEFT, padx=(10, 2))

        self._cnt = ttk.Label(bot, text="", foreground=GRAY)
        self._cnt.pack(side=tk.RIGHT)

    def _on_dbl_click(self) -> None:
        self._edit()

    def _reset(self) -> None:
        self._v_search.set("")
        self._v_status.set("")
        self._v_prio.set("")
        self._v_overdue.set(False)
        self.refresh()

    def refresh(self) -> None:
        try:
            self._clients_cache = {c["id"]: c["name"]
                                   for c in self.api.get("/clients", limit=500)}
            self._deals_cache   = {d["id"]: d["title"]
                                   for d in self.api.get("/deals", limit=500)}
        except Exception:
            pass
        try:
            overdue = self._v_overdue.get() or None   # True | None
            data = self.api.get("/tasks",
                                search=self._v_search.get().strip() or None,
                                status=self._v_status.get() or None,
                                priority=self._v_prio.get() or None,
                                overdue_only=overdue,
                                limit=500)
        except Exception as e:
            self._err(e)
            return
        self._fill_tree(data, lambda r: (
            r["id"],
            r.get("title", ""),
            _lbl(r.get("priority", "")),
            _lbl(r.get("status", "")),
            r.get("due_date", "") or "—",
            self._clients_cache.get(r.get("client_id"), "—"),
            self._deals_cache.get(r.get("deal_id"), "—"),
            _fmt_dt(r.get("updated_at", "")),
        ))
        self._count_label(self._cnt, len(data))

    # ── CRUD ──────────────────────────────────────────────────────────────────

    def _create(self) -> None:
        if TaskDialog(self, self.api).result is not None:
            self.refresh()

    def _edit(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите задачу в списке.", parent=self)
            return
        try:
            item = self.api.get(f"/tasks/{iid}")
        except APIError as e:
            self._err(e)
            return
        if TaskDialog(self, self.api, existing=item).result is not None:
            self.refresh()

    def _complete(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите задачу.", parent=self)
            return
        try:
            self.api.patch(f"/tasks/{iid}/complete")
            self.refresh()
        except APIError as e:
            self._err(e)

    def _change_status(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите задачу.", parent=self)
            return
        try:
            item = self.api.get(f"/tasks/{iid}")
        except APIError as e:
            self._err(e)
            return
        new_status = _ask_status(self, "Статус задачи", TASK_STATUSES,
                                  current=item.get("status", ""))
        if new_status:
            try:
                self.api.patch(f"/tasks/{iid}/status", status=new_status)
                self.refresh()
            except APIError as e:
                self._err(e)

    def _delete(self) -> None:
        iid = self._selected_id()
        if not iid:
            messagebox.showinfo("Выбор", "Выберите задачу.", parent=self)
            return
        if self._confirm("Удалить задачу безвозвратно?"):
            try:
                self.api.delete(f"/tasks/{iid}")
                self.refresh()
            except APIError as e:
                self._err(e)

    def _export_report(self) -> None:
        rows = [self.tv.item(iid)["values"] for iid in self.tv.get_children()]
        if not rows:
            messagebox.showinfo("Экспорт", "Нет данных для экспорта.", parent=self)
            return
        today = datetime.now().strftime("%d.%m.%Y")
        # Сводная информация из дашборда (раздел «Задачи»)
        summary: List[tuple] = []
        try:
            s  = self.api.get("/dashboard")
            st = s.get("tasks", {})
            summary = [
                ("Ожидают выполнения", st.get("pending",   0)),
                ("Выполнено",          st.get("done",      0)),
                ("Просрочено",         st.get("overdue",   0)),
                ("Высокий приоритет",  st.get("high_prio", 0)),
            ]
        except Exception:
            pass
        _export_to_gsheet(self, f"CRM — Задачи {today}", list(self.HDRS), rows,
                          summary=summary)


# ══════════════════════════════════════════════════════════════════════════════
#  Главное приложение
# ══════════════════════════════════════════════════════════════════════════════

class App:

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Mini CRM")
        self.root.geometry("1150x660")
        self.root.minsize(900, 560)
        self.root.configure(bg=BG)
        self.api = API()
        self._apply_style()
        self._build()

    def _apply_style(self) -> None:
        s = ttk.Style()
        s.theme_use("clam")
        for w in ("TFrame", "TLabel", "TLabelframe", "TLabelframe.Label",
                  "TCheckbutton", "TNotebook", "TNotebook.Tab"):
            try:
                s.configure(w, background=BG)
            except Exception:
                pass
        s.configure("TButton", padding=(6, 3))
        s.configure("Treeview", rowheight=24, fieldbackground=WHITE)
        s.configure("Treeview.Heading", font=("TkDefaultFont", 9, "bold"))
        s.map("Treeview", background=[("selected", BLUE)],
              foreground=[("selected", WHITE)])

    def _build(self) -> None:
        # ── шапка ────────────────────────────────────────────────────────────
        header = tk.Frame(self.root, bg=DARK, height=46)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        tk.Label(header, text="Mini CRM", fg=WHITE, bg=DARK,
                 font=("TkDefaultFont", 13, "bold")).pack(side=tk.LEFT, padx=16, pady=8)

        self._conn_lbl = tk.Label(header, text="", fg=WHITE, bg=DARK,
                                   font=("TkDefaultFont", 9))
        self._conn_lbl.pack(side=tk.RIGHT, padx=16)

        tk.Button(
            header, text="⚙ Настройки Google",
            fg=WHITE, bg="#1A5276", activebackground="#154360",
            activeforeground=WHITE, bd=0, padx=10, pady=4, cursor="hand2",
            font=("TkDefaultFont", 9),
            command=lambda: GoogleSettingsDialog(self.root),
        ).pack(side=tk.RIGHT, padx=(0, 6), pady=6)

        # ── вкладки ───────────────────────────────────────────────────────────
        nb = ttk.Notebook(self.root)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=(6, 0))

        self._dashboard = DashboardFrame(nb, self.api)
        self._clients   = ClientsFrame(nb, self.api)
        self._deals     = DealsFrame(nb, self.api)
        self._tasks     = TasksFrame(nb, self.api)

        nb.add(self._dashboard, text="  Дашборд   ")
        nb.add(self._clients,   text="  Клиенты   ")
        nb.add(self._deals,     text="  Сделки    ")
        nb.add(self._tasks,     text="  Задачи    ")

        nb.bind("<<NotebookTabChanged>>", self._on_tab)

        # ── строка состояния ──────────────────────────────────────────────────
        status_bar = tk.Frame(self.root, bg="#DDE1E7", height=22)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        status_bar.pack_propagate(False)
        self._status_lbl = tk.Label(status_bar, text="", bg="#DDE1E7",
                                     font=("TkDefaultFont", 8), anchor=tk.W)
        self._status_lbl.pack(side=tk.LEFT, padx=10)

        self._tabs = [self._dashboard, self._clients, self._deals, self._tasks]
        self._nb   = nb
        self._ping()

    def _on_tab(self, event: tk.Event) -> None:
        idx = self._nb.index(self._nb.select())
        tab = self._tabs[idx]
        if hasattr(tab, "refresh"):
            tab.refresh()

    def _ping(self) -> None:
        ok = self.api.ping()
        if ok:
            self._conn_lbl.config(text="[online]  http://127.0.0.1:8000")
            self._status_lbl.config(
                text="Сервер: подключено  |  Swagger UI: http://127.0.0.1:8000/docs",
                fg=GREEN)
        else:
            self._conn_lbl.config(text="[offline]")
            self._status_lbl.config(
                text="Сервер недоступен — запустите:  venv\\Scripts\\python.exe crm\\crm.py",
                fg=RED)
        self.root.after(10_000, self._ping)

    def run(self) -> None:
        self.root.mainloop()


# ══════════════════════════════════════════════════════════════════════════════
#  Точка входа
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    App().run()
