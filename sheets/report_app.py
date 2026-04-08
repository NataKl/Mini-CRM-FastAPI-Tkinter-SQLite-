"""
report_app.py
=============
Tkinter-приложение для генерации случайных отчётов по отделам
и записи их в Google Таблицу в виде красиво оформленного документа.

Запуск:
    python report_app.py

Требования:
    pip install -r requirements.txt
    .env с GOOGLE_SERVICE_ACCOUNT_FILE и GOOGLE_SPREADSHEET_ID
"""

from __future__ import annotations

import os
import random
import threading
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import messagebox, ttk

# ── Загрузка .env из корня проекта ────────────────────────────────────────────
try:
    from dotenv import load_dotenv

    load_dotenv(dotenv_path=Path(__file__).with_name(".env"), override=False)
except Exception:
    pass

try:
    from tkcalendar import DateEntry

    _HAS_CAL = True
except ImportError:
    _HAS_CAL = False

from gsheet_api import GSheetClient  # noqa: E402


# ══════════════════════════════════════════════════════════════════════════════
#  Палитра Google Sheets
# ══════════════════════════════════════════════════════════════════════════════

def _rgb(r: int, g: int, b: int) -> Dict[str, float]:
    return {"red": r / 255.0, "green": g / 255.0, "blue": b / 255.0}


C: Dict[str, Dict[str, float]] = {
    "title_bg":    _rgb(31,  73, 125),   # тёмно-синий — заголовок отчёта
    "title_fg":    _rgb(255, 255, 255),
    "company_bg":  _rgb(68, 114, 196),   # средне-синий — строка организации
    "company_fg":  _rgb(255, 255, 255),
    "ilabel_bg":   _rgb(197, 217, 241),  # светло-голубой — подписи реквизитов
    "ilabel_fg":   _rgb(31,  73, 125),
    "ivalue_bg":   _rgb(255, 255, 255),  # белый — значения реквизитов
    "ivalue_fg":   _rgb(30,  30,  30),
    "thead_bg":    _rgb(47,  84, 150),   # шапка таблицы
    "thead_fg":    _rgb(255, 255, 255),
    "row_odd":     _rgb(221, 231, 242),  # нечётная строка данных
    "row_even":    _rgb(255, 255, 255),  # чётная строка данных
    "total_bg":    _rgb(189, 215, 238),  # итого
    "total_fg":    _rgb(31,  73, 125),
    "sep_bg":      _rgb(255, 255, 255),  # разделитель
    "footer_bg":   _rgb(242, 242, 242),  # подписи
    "footer_fg":   _rgb(89,  89,  89),
    "border":      _rgb(68, 114, 196),
    "border_hdr":  _rgb(31,  73, 125),
    "black":       _rgb(0,   0,   0),
    "white":       _rgb(255, 255, 255),
}


# ══════════════════════════════════════════════════════════════════════════════
#  Данные для случайного отчёта
#  Каждый показатель: (название, единица, факт_мин, факт_макс, план_мин, план_макс)
# ══════════════════════════════════════════════════════════════════════════════

DEPT_METRICS: Dict[str, List[Tuple]] = {
    "IT-отдел": [
        ("Закрытые тикеты / задачи",         "шт.",       120, 160,  90, 140),
        ("Открытые инциденты",                "шт.",         2,  12,   5,  18),
        ("Среднее время отклика на заявки",   "ч.",          1,   4,   2,   6),
        ("Доступность систем (uptime)",       "%",         99.2, 99.9, 99.0, 99.8),
        ("Выполнено задач в срок",            "%",          85,  99,  80,  95),
        ("Критических инцидентов",            "шт.",         0,   3,   0,   5),
        ("Обновлений ПО / систем",            "шт.",        10,  40,   8,  35),
        ("Сотрудников в отделе",              "чел.",        8,  15,   8,  15),
        ("Прошли плановое обучение",          "%",          80, 100,  70, 100),
        ("Удовлетворённость пользователей",   "%",          88,  98,  85,  95),
        ("Новых автоматизаций",               "шт.",         2,  10,   1,   8),
        ("Часов технических работ (план.)",   "ч.",          4,  40,   4,  50),
    ],
    "Отдел продаж": [
        ("Закрытых сделок",                   "шт.",        30,  80,  25,  70),
        ("Новых клиентов",                    "чел.",       10,  50,   8,  45),
        ("Выручка",                           "тыс. руб.", 800, 3500, 700, 3200),
        ("Средний чек",                       "тыс. руб.",  20,  80,  18,  75),
        ("Конверсия лидов",                   "%",          15,  40,  12,  35),
        ("Исходящих звонков",                 "шт.",       200, 600, 180, 550),
        ("Проведено встреч / демо",           "шт.",        20, 100,  15,  90),
        ("Выполнение плана продаж",           "%",          85, 115,  90, 110),
        ("Индекс NPS",                        "балл",       7.0, 9.5, 7.0, 9.0),
        ("Сотрудников в отделе",              "чел.",        5,  20,   5,  20),
        ("Повторные продажи",                 "%",          30,  70,  25,  65),
        ("Лидов в активной воронке",          "шт.",       100, 400,  90, 380),
    ],
    "HR-отдел": [
        ("Открытых вакансий",                 "шт.",         5,  25,   4,  22),
        ("Закрытых вакансий",                 "шт.",         3,  18,   3,  15),
        ("Среднее время закрытия вакансии",   "дн.",        15,  45,  20,  50),
        ("Текучесть кадров",                  "%",           3,  12,   2,  10),
        ("Проведено собеседований",           "шт.",        20,  80,  15,  70),
        ("Охвачено обучением",                "чел.",       10,  50,   8,  45),
        ("Удовлетворённость персонала",       "%",          75,  95,  70,  90),
        ("Дней нетрудоспособности (всего)",   "дн.",         5,  30,   0,  25),
        ("Штатная численность (факт)",        "чел.",       80, 150,  85, 155),
        ("Выполнение KPI персонала",          "%",          78,  98,  75,  95),
        ("Адаптировано новых сотрудников",    "чел.",        3,  15,   3,  12),
        ("Взыскания / нарушения дисциплины",  "шт.",         0,   5,   0,   3),
    ],
    "Финансовый отдел": [
        ("Обработано счетов / поручений",     "шт.",        50, 200,  45, 180),
        ("Дебиторская задолженность",         "тыс. руб.", 200, 1500, 180, 1400),
        ("Кредиторская задолженность",        "тыс. руб.", 100, 800,  90, 750),
        ("Исполнение бюджета",                "%",          90, 102,  95, 100),
        ("Выручка компании",                  "млн. руб.",   5,  20, 4.5,  18),
        ("Чистая прибыль",                    "млн. руб.",   1,   8, 0.9,   7),
        ("Рентабельность продаж",             "%",           5,  30,   5,  28),
        ("Налоговая нагрузка",                "%",          15,  25,  15,  22),
        ("Ошибок в финансовой отчётности",    "шт.",         0,   5,   0,   3),
        ("Срок закрытия отчётного периода",   "дн.",         3,   8,   2,   7),
        ("Просроченная дебиторская задолж.",  "%",           2,  20,   0,  15),
        ("Прогноз исполнения бюджета",        "%",          88, 105,  90, 100),
    ],
    "Маркетинг": [
        ("Охват рекламных кампаний",          "тыс. чел.",  10, 500,   8, 450),
        ("Лидов сгенерировано",               "шт.",       100, 800,  90, 700),
        ("Конверсия сайта",                   "%",           2,   8,   2,   7),
        ("Стоимость лида (CPL)",              "руб.",       50, 500,  60, 600),
        ("ROAS",                              "х",           2,   8,   2,   7),
        ("Публикаций в соцсетях",             "шт.",        20, 120,  15, 100),
        ("Email-рассылок отправлено",         "тыс.",        5,  50,   4,  45),
        ("Open rate рассылок",                "%",          15,  40,  15,  35),
        ("Освоение рекламного бюджета",       "%",          85, 105,  90, 100),
        ("Упоминаний бренда (медиа)",         "шт.",        30, 300,  25, 280),
        ("CTR рекламных объявлений",          "%",         1.5,   6, 1.2,   5),
        ("Органический трафик на сайт",       "тыс. виз.",   5, 100,   4,  90),
    ],
    "Производство": [
        ("Объём выпущенной продукции",        "шт.",       500, 3000, 450, 2800),
        ("Процент брака",                     "%",           0,   5,   0,   4),
        ("Загрузка производственных мощностей", "%",        70,  98,  75,  95),
        ("Простои оборудования",              "ч.",          0,  40,   0,  30),
        ("Выполнение производственного плана","% ",         88, 105,  90, 100),
        ("Расход сырья",                      "тонн",      100, 500,  95, 490),
        ("Энергопотребление",                 "тыс. кВт·ч",  5,  20, 4.5,  18),
        ("Аварий / инцидентов",               "шт.",         0,   3,   0,   2),
        ("Выработка на одного сотрудника",    "шт./чел.",   50, 200,  45, 190),
        ("Себестоимость единицы продукции",   "руб.",      100, 1000, 110, 1050),
        ("Сдано ОТК с 1-го предъявления",     "%",          85,  99,  80,  98),
        ("Рекламаций от клиентов",            "шт.",         0,   8,   0,   6),
    ],
    "Логистика": [
        ("Выполнено доставок",                "шт.",       200, 1000, 180, 950),
        ("Своевременность доставок",          "%",          90,  99,  88,  98),
        ("Средний срок доставки",             "дн.",         1,   5,   1,   6),
        ("Складской остаток",                 "ед.",        500, 5000, 400, 4800),
        ("Оборачиваемость склада",            "раз/мес.",    2,  10,   2,   9),
        ("Потери / недостачи",                "тыс. руб.",   0,  50,   0,  40),
        ("Стоимость логистики",               "тыс. руб.",  50, 500,  45, 480),
        ("Возвраты",                          "%",           1,   8,   0,   6),
        ("Загрузка транспортных средств",     "%",          60,  95,  65,  90),
        ("Рекламаций по доставке",            "шт.",         0,  20,   0,  15),
        ("Обработано заказов на складе",      "шт.",        300, 2000, 280, 1900),
        ("Ошибок при комплектации",           "шт.",         0,  30,   0,  25),
    ],
    "Юридический отдел": [
        ("Договоров подготовлено",            "шт.",        20, 100,  18,  90),
        ("Договоров согласовано",             "шт.",        18,  95,  16,  88),
        ("Судебных дел в производстве",       "шт.",         0,  15,   0,  12),
        ("Судебных дел выиграно",             "шт.",         0,  10,   0,  10),
        ("Претензий получено",                "шт.",         2,  20,   2,  18),
        ("Претензий урегулировано",           "%",          70,  99,  75,  95),
        ("Регуляторных нарушений выявлено",   "шт.",         0,   3,   0,   2),
        ("Проверок пройдено",                 "шт.",         0,   5,   0,   4),
        ("Юридических консультаций дано",     "шт.",        30, 150,  25, 140),
        ("Выполнение внутренних регламентов", "%",          90, 100,  88, 100),
        ("Регистраций интеллект. собств.",    "шт.",         0,   5,   0,   4),
        ("Комплаенс-нарушений",               "шт.",         0,   2,   0,   1),
    ],
}

NOTES_POOL: List[str] = [
    "Выполнено в полном объёме",
    "Требует дополнительного контроля",
    "Показатель в норме",
    "Плановые значения скорректированы",
    "Отклонение в допустимых пределах",
    "Необходим анализ причин отклонения",
    "Перевыполнение плана",
    "Под контролем руководства",
    "В работе, срок не истёк",
    "Решение принято",
    "", "", "", "",  # пустые чаще встречаются
]


def _gen_rows(department: str, n: int = 12) -> List[Dict[str, Any]]:
    """Генерирует n случайных строк метрик для указанного отдела."""
    metrics = DEPT_METRICS.get(department, list(DEPT_METRICS.values())[0])
    chosen = random.sample(metrics, min(n, len(metrics)))
    result: List[Dict[str, Any]] = []
    for i, (name, unit, f_lo, f_hi, p_lo, p_hi) in enumerate(chosen, start=1):
        is_float = isinstance(f_lo, float) or isinstance(p_lo, float)
        if is_float:
            fact: Any = round(random.uniform(f_lo, f_hi), 2)
            plan: Any = round(random.uniform(p_lo, p_hi), 2)
            dev: Any = round(fact - plan, 2)
        else:
            fact = random.randint(int(f_lo), int(f_hi))
            plan = random.randint(int(p_lo), int(p_hi))
            dev = fact - plan
        pct = round((fact / plan * 100) if plan else 0.0, 1)
        result.append({
            "num": i,
            "name": name,
            "unit": unit,
            "plan": plan,
            "fact": fact,
            "dev": dev,
            "pct": pct,
            "note": random.choice(NOTES_POOL),
        })
    return result


# ══════════════════════════════════════════════════════════════════════════════
#  Построение отчёта в Google Sheets
# ══════════════════════════════════════════════════════════════════════════════

# Столбцы (0-based): A=0 … H=7
_NCOLS = 8
_COL_WIDTHS = [50, 270, 85, 115, 115, 115, 105, 185]  # px

# Заголовки столбцов таблицы
_TABLE_HEADERS = ["№", "Показатель", "Ед. изм.", "План", "Факт", "Отклонение", "% выполн.", "Примечания"]


def _solid_border(color: Optional[Dict] = None, width: int = 1, style: str = "SOLID") -> Dict:
    return {"style": style, "width": width, "color": color or C["border"]}


def _all_borders(thick: bool = False) -> Dict:
    s = _solid_border(C["border_hdr"], 2, "SOLID_MEDIUM") if thick else _solid_border(C["border"])
    return {"top": s, "bottom": s, "left": s, "right": s}


def _fmt_req(
    sheet_id: int,
    r1: int, c1: int, r2: int, c2: int,
    *,
    bg: Optional[Dict] = None,
    fg: Optional[Dict] = None,
    bold: bool = False,
    italic: bool = False,
    size: int = 10,
    h_align: str = "LEFT",
    v_align: str = "MIDDLE",
    wrap: bool = True,
    borders: Optional[Dict] = None,
) -> Dict:
    """Возвращает repeatCell-запрос для форматирования диапазона."""
    fmt: Dict[str, Any] = {
        "textFormat": {
            "foregroundColor": fg if fg else C["black"],
            "bold": bold,
            "italic": italic,
            "fontSize": size,
        },
        "horizontalAlignment": h_align,
        "verticalAlignment": v_align,
        "wrapStrategy": "WRAP" if wrap else "CLIP",
    }
    if bg:
        fmt["backgroundColor"] = bg
    if borders:
        fmt["borders"] = borders
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": c1,
                "endColumnIndex": c2,
            },
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat",
        }
    }


def _merge_req(sheet_id: int, r1: int, c1: int, r2: int, c2: int) -> Dict:
    return {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2,
            },
            "mergeType": "MERGE_ALL",
        }
    }


def _col_width_req(sheet_id: int, col: int, px: int) -> Dict:
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": col, "endIndex": col + 1},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def _row_height_req(sheet_id: int, row: int, px: int, end_row: Optional[int] = None) -> Dict:
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": row, "endIndex": (end_row or row + 1)},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def build_report_in_sheet(
    client: GSheetClient,
    company: str,
    department: str,
    responsible: str,
    date_from: date,
    date_to: date,
    sheet_title: str,
) -> str:
    """
    Создаёт новый лист, записывает в него сгенерированный отчёт
    и применяет полное форматирование «в стиле Word-документа».
    Возвращает итоговое название листа.
    """
    # ── 1. Создаём лист ───────────────────────────────────────────────────────
    resp = client.create_sheet(sheet_title, rows=200, cols=_NCOLS)
    sheet_id: int = resp["replies"][0]["addSheet"]["properties"]["sheetId"]

    # ── 2. Генерируем данные ─────────────────────────────────────────────────
    data_rows = _gen_rows(department, n=12)
    n = len(data_rows)

    today_str = date.today().strftime("%d.%m.%Y")
    period_str = (
        f"{date_from.strftime('%d.%m.%Y')} — {date_to.strftime('%d.%m.%Y')}"
    )

    # Индексы строк (0-based)
    R_TITLE = 0        # «ОТЧЁТ О ДЕЯТЕЛЬНОСТИ ОТДЕЛА»
    R_COMPANY = 1      # название организации
    R_SEP1 = 2         # пустой разделитель
    R_PER = 3          # Период:
    R_DEPT = 4         # Отдел:
    R_RESP = 5         # Ответственный:
    R_DATE = 6         # Дата составления:
    R_SEP2 = 7         # пустой разделитель
    R_THEAD = 8        # шапка таблицы
    R_DATA0 = 9        # первая строка данных
    R_DATA1 = 9 + n - 1  # последняя строка данных
    R_TOTAL = 9 + n    # ИТОГО
    R_SEP3 = 10 + n    # пустой разделитель
    R_SIGN1 = 11 + n   # Отчёт составил
    R_SIGN2 = 12 + n   # Утверждено
    R_SIGN3 = 13 + n   # Дата подписания

    # ── 3. Формируем значения ────────────────────────────────────────────────
    plan_sum = round(sum(r["plan"] for r in data_rows), 2)
    fact_sum = round(sum(r["fact"] for r in data_rows), 2)
    dev_sum  = round(fact_sum - plan_sum, 2)
    pct_avg  = round(sum(r["pct"] for r in data_rows) / n, 1)

    values: List[List[Any]] = []
    values.append(["ОТЧЁТ О ДЕЯТЕЛЬНОСТИ ОТДЕЛА"] + [""] * 7)   # R_TITLE
    values.append([company] + [""] * 7)                          # R_COMPANY
    values.append([""] * _NCOLS)                                 # R_SEP1
    values.append(["Период:",            period_str] + [""] * 6) # R_PER
    values.append(["Отдел:",             department] + [""] * 6) # R_DEPT
    values.append(["Ответственный:",     responsible] + [""] * 6)# R_RESP
    values.append(["Дата составления:",  today_str] + [""] * 6)  # R_DATE
    values.append([""] * _NCOLS)                                 # R_SEP2
    values.append(_TABLE_HEADERS)                                # R_THEAD

    for r in data_rows:
        dev_str = (f"+{r['dev']}" if isinstance(r["dev"], float) and r["dev"] > 0
                   else (f"+{r['dev']}" if isinstance(r["dev"], int) and r["dev"] > 0
                         else str(r["dev"])))
        values.append([
            r["num"], r["name"], r["unit"],
            r["plan"], r["fact"], dev_str,
            f"{r['pct']}%", r["note"],
        ])

    dev_sum_str = (f"+{dev_sum}" if dev_sum > 0 else str(dev_sum))
    values.append(["", "ИТОГО / СРЕДНЕЕ", "", plan_sum, fact_sum,
                   dev_sum_str, f"{pct_avg}%", ""])  # R_TOTAL
    values.append([""] * _NCOLS)                    # R_SEP3
    values.append([
        "Отчёт составил:", "____________", "", f"/ {responsible} /",
        "", "", "", "",
    ])
    values.append(["Утверждено:",  "____________"] + [""] * 6)
    values.append([f"Дата:  «___» ____________ {date.today().year} г."] + [""] * 7)

    # Записываем значения (с правильным экранированием названия листа)
    safe_title = sheet_title.replace("'", "\\'")
    client.update_range(f"'{safe_title}'!A1", values)

    # ── 4. Формируем запросы на форматирование ───────────────────────────────
    reqs: List[Dict] = []

    # ---- Объединение ячеек ----
    # Заголовок и компания — на всю ширину
    reqs += [
        _merge_req(sheet_id, R_TITLE,   0, R_TITLE + 1,   _NCOLS),
        _merge_req(sheet_id, R_COMPANY, 0, R_COMPANY + 1, _NCOLS),
        _merge_req(sheet_id, R_SEP1,    0, R_SEP1 + 1,    _NCOLS),
        _merge_req(sheet_id, R_SEP2,    0, R_SEP2 + 1,    _NCOLS),
        _merge_req(sheet_id, R_SEP3,    0, R_SEP3 + 1,    _NCOLS),
        _merge_req(sheet_id, R_SIGN3,   0, R_SIGN3 + 1,   _NCOLS),
    ]
    # Реквизитные строки: B:H объединены
    for r in (R_PER, R_DEPT, R_RESP, R_DATE):
        reqs.append(_merge_req(sheet_id, r, 1, r + 1, _NCOLS))
    # В строке ИТОГО: B:C объединены (надпись «ИТОГО / СРЕДНЕЕ»)
    reqs.append(_merge_req(sheet_id, R_TOTAL, 1, R_TOTAL + 1, 3))
    # Строки подписей
    for c_start, c_end in [(1, 3), (3, 6)]:
        reqs.append(_merge_req(sheet_id, R_SIGN1, c_start, R_SIGN1 + 1, c_end))
    reqs.append(_merge_req(sheet_id, R_SIGN2, 1, R_SIGN2 + 1, 4))

    # ---- Форматирование строк ----

    # Заголовок отчёта
    reqs.append(_fmt_req(sheet_id, R_TITLE, 0, R_TITLE + 1, _NCOLS,
                         bg=C["title_bg"], fg=C["title_fg"],
                         bold=True, size=14, h_align="CENTER"))

    # Строка организации
    reqs.append(_fmt_req(sheet_id, R_COMPANY, 0, R_COMPANY + 1, _NCOLS,
                         bg=C["company_bg"], fg=C["company_fg"],
                         bold=True, size=11, h_align="CENTER"))

    # Разделители
    for r in (R_SEP1, R_SEP2, R_SEP3):
        reqs.append(_fmt_req(sheet_id, r, 0, r + 1, _NCOLS, bg=C["sep_bg"]))

    # Подписи реквизитов (столбец A)
    reqs.append(_fmt_req(sheet_id, R_PER, 0, R_DATE + 1, 1,
                         bg=C["ilabel_bg"], fg=C["ilabel_fg"], bold=True))

    # Значения реквизитов (столбцы B:H)
    reqs.append(_fmt_req(sheet_id, R_PER, 1, R_DATE + 1, _NCOLS,
                         bg=C["ivalue_bg"], fg=C["ivalue_fg"]))

    # Шапка таблицы
    reqs.append(_fmt_req(sheet_id, R_THEAD, 0, R_THEAD + 1, _NCOLS,
                         bg=C["thead_bg"], fg=C["thead_fg"],
                         bold=True, size=10, h_align="CENTER",
                         borders=_all_borders(thick=True)))

    # Строки данных (чередующаяся заливка)
    for i, ri in enumerate(range(R_DATA0, R_DATA1 + 1)):
        bg = C["row_odd"] if i % 2 == 0 else C["row_even"]
        reqs.append(_fmt_req(sheet_id, ri, 0, ri + 1, _NCOLS,
                             bg=bg, size=10, borders=_all_borders()))
        # № — по центру
        reqs.append(_fmt_req(sheet_id, ri, 0, ri + 1, 1, h_align="CENTER"))
        # Числовые столбцы (D-G) — по правому краю
        reqs.append(_fmt_req(sheet_id, ri, 3, ri + 1, 7, h_align="RIGHT"))

    # Строка ИТОГО
    reqs.append(_fmt_req(sheet_id, R_TOTAL, 0, R_TOTAL + 1, _NCOLS,
                         bg=C["total_bg"], fg=C["total_fg"],
                         bold=True, size=10, borders=_all_borders(thick=True)))
    reqs.append(_fmt_req(sheet_id, R_TOTAL, 3, R_TOTAL + 1, 7, h_align="RIGHT"))

    # Подписи (footer)
    for r in (R_SIGN1, R_SIGN2, R_SIGN3):
        reqs.append(_fmt_req(sheet_id, r, 0, r + 1, _NCOLS,
                             bg=C["footer_bg"], fg=C["footer_fg"],
                             italic=True, size=10))

    # ---- Ширина столбцов ----
    for col, px in enumerate(_COL_WIDTHS):
        reqs.append(_col_width_req(sheet_id, col, px))

    # ---- Высота строк ----
    row_heights = {
        R_TITLE: 55, R_COMPANY: 40,
        R_SEP1: 10,  R_SEP2: 10, R_SEP3: 18,
        R_THEAD: 44, R_TOTAL: 38,
        R_SIGN1: 30, R_SIGN2: 30, R_SIGN3: 30,
    }
    for r, h in row_heights.items():
        reqs.append(_row_height_req(sheet_id, r, h))

    # Реквизитные строки
    for r in (R_PER, R_DEPT, R_RESP, R_DATE):
        reqs.append(_row_height_req(sheet_id, r, 28))

    # Строки данных
    reqs.append(_row_height_req(sheet_id, R_DATA0, 30, end_row=R_DATA1 + 1))

    # ── 5. Применяем всё форматирование одним запросом ───────────────────────
    client.apply_requests(reqs)

    return sheet_title


# ══════════════════════════════════════════════════════════════════════════════
#  Tkinter-приложение
# ══════════════════════════════════════════════════════════════════════════════

class ReportApp(tk.Tk):
    """Основное окно приложения."""

    DEPARTMENTS = list(DEPT_METRICS.keys())

    # ── Цвета UI ──
    BG         = "#f0f4f8"
    HEADER_BG  = "#1e3a5f"
    HEADER_FG  = "#ffffff"
    FRAME_BG   = "#ffffff"
    LABEL_FG   = "#1e293b"
    HINT_FG    = "#64748b"
    BTN_BG     = "#2563eb"
    BTN_FG     = "#ffffff"
    BTN_ACT    = "#1d4ed8"
    OK_FG      = "#16a34a"
    ERR_FG     = "#dc2626"
    WARN_FG    = "#d97706"

    def __init__(self) -> None:
        super().__init__()
        self.title("Генератор отчётов → Google Sheets")
        self.geometry("590x570")
        self.resizable(False, False)
        self.configure(bg=self.BG)

        self._client: Optional[GSheetClient] = None
        self._busy = False

        self._build_header()
        self._build_form()
        self._build_button()
        self._build_status()
        self._init_client()

    # ── Шапка ─────────────────────────────────────────────────────────────────
    def _build_header(self) -> None:
        hdr = tk.Frame(self, bg=self.HEADER_BG, pady=18)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="📊  Генератор отчётов",
            bg=self.HEADER_BG, fg=self.HEADER_FG,
            font=("Segoe UI", 16, "bold"),
        ).pack()
        tk.Label(
            hdr, text="Автоматическая запись в Google Sheets",
            bg=self.HEADER_BG, fg="#93c5fd",
            font=("Segoe UI", 9),
        ).pack()

    # ── Форма ─────────────────────────────────────────────────────────────────
    def _build_form(self) -> None:
        wrap = tk.Frame(self, bg=self.BG, padx=20, pady=14)
        wrap.pack(fill="x")

        frm = tk.Frame(
            wrap, bg=self.FRAME_BG,
            highlightbackground="#cbd5e1",
            highlightthickness=1,
        )
        frm.pack(fill="x")
        frm.columnconfigure(1, weight=1)

        def label(text: str, row: int, hint: bool = False) -> None:
            tk.Label(
                frm, text=text,
                bg=self.FRAME_BG,
                fg=self.HINT_FG if hint else self.LABEL_FG,
                font=("Segoe UI", 8 if hint else 9),
                anchor="w",
            ).grid(row=row, column=0, columnspan=2,
                   sticky="w", padx=18, pady=(10 if not hint else 0, 2))

        def entry(var: tk.StringVar, row: int) -> tk.Entry:
            e = tk.Entry(
                frm, textvariable=var, font=("Segoe UI", 10),
                relief="solid", bd=1, highlightthickness=0,
            )
            e.grid(row=row, column=0, columnspan=2,
                   sticky="ew", padx=18, pady=(0, 4), ipady=4)
            return e

        # Организация
        label("Организация", 0)
        self._company_var = tk.StringVar(value="ООО «Ромашка»")
        entry(self._company_var, 1)

        # Отдел
        label("Отдел", 2)
        self._dept_var = tk.StringVar(value=self.DEPARTMENTS[0])
        cb = ttk.Combobox(
            frm, textvariable=self._dept_var,
            values=self.DEPARTMENTS, state="readonly",
            font=("Segoe UI", 10),
        )
        cb.grid(row=3, column=0, columnspan=2,
                sticky="ew", padx=18, pady=(0, 4), ipady=4)

        # Ответственный
        label("Ответственный (ФИО)", 4)
        self._resp_var = tk.StringVar(value="Иванов Иван Иванович")
        entry(self._resp_var, 5)

        # Период
        label("Период отчёта", 6)
        date_row = tk.Frame(frm, bg=self.FRAME_BG)
        date_row.grid(row=7, column=0, columnspan=2,
                      sticky="ew", padx=18, pady=(0, 4))

        first_of_month = date.today().replace(day=1)

        tk.Label(date_row, text="с", bg=self.FRAME_BG, fg=self.LABEL_FG,
                 font=("Segoe UI", 10)).pack(side="left", padx=(0, 6))

        if _HAS_CAL:
            self._date_from = DateEntry(
                date_row, date_pattern="dd.mm.yyyy",
                width=12, font=("Segoe UI", 10),
                background="#1e3a5f", foreground="white",
            )
            self._date_from.set_date(first_of_month)
            self._date_from.pack(side="left")

            tk.Label(date_row, text="по", bg=self.FRAME_BG, fg=self.LABEL_FG,
                     font=("Segoe UI", 10)).pack(side="left", padx=(12, 6))

            self._date_to = DateEntry(
                date_row, date_pattern="dd.mm.yyyy",
                width=12, font=("Segoe UI", 10),
                background="#1e3a5f", foreground="white",
            )
            self._date_to.set_date(date.today())
            self._date_to.pack(side="left")
        else:
            self._date_from_var = tk.StringVar(
                value=first_of_month.strftime("%d.%m.%Y")
            )
            tk.Entry(date_row, textvariable=self._date_from_var,
                     width=12, font=("Segoe UI", 10),
                     relief="solid", bd=1).pack(side="left")

            tk.Label(date_row, text="по", bg=self.FRAME_BG, fg=self.LABEL_FG,
                     font=("Segoe UI", 10)).pack(side="left", padx=(12, 6))

            self._date_to_var = tk.StringVar(
                value=date.today().strftime("%d.%m.%Y")
            )
            tk.Entry(date_row, textvariable=self._date_to_var,
                     width=12, font=("Segoe UI", 10),
                     relief="solid", bd=1).pack(side="left")

            tk.Label(
                date_row, text="  формат дд.мм.гггг",
                bg=self.FRAME_BG, fg=self.HINT_FG,
                font=("Segoe UI", 8),
            ).pack(side="left", padx=(8, 0))

        # Название листа
        label("Название листа в Google Sheets (оставьте пустым — сгенерируется автоматически)", 8, hint=True)
        self._sheet_name_var = tk.StringVar(value="")
        entry(self._sheet_name_var, 9)

        # нижний отступ
        tk.Frame(frm, bg=self.FRAME_BG, height=6).grid(row=10, column=0)

    # ── Кнопка ────────────────────────────────────────────────────────────────
    def _build_button(self) -> None:
        btn_wrap = tk.Frame(self, bg=self.BG, pady=14)
        btn_wrap.pack(fill="x", padx=20)

        self._btn = tk.Button(
            btn_wrap,
            text="🚀  Сформировать отчёт",
            font=("Segoe UI", 12, "bold"),
            bg=self.BTN_BG, fg=self.BTN_FG,
            activebackground=self.BTN_ACT, activeforeground=self.BTN_FG,
            relief="flat", bd=0, cursor="hand2",
            pady=10,
            command=self._on_generate,
        )
        self._btn.pack(fill="x")

    # ── Статусная строка ──────────────────────────────────────────────────────
    def _build_status(self) -> None:
        sep = tk.Frame(self, bg="#cbd5e1", height=1)
        sep.pack(fill="x", padx=20)

        status_wrap = tk.Frame(self, bg=self.BG, padx=20, pady=10)
        status_wrap.pack(fill="x")

        self._status_var = tk.StringVar(value="⏳ Инициализация…")
        self._status_lbl = tk.Label(
            status_wrap,
            textvariable=self._status_var,
            bg=self.BG, fg=self.HINT_FG,
            font=("Segoe UI", 9),
            anchor="w", wraplength=540, justify="left",
        )
        self._status_lbl.pack(fill="x")

    # ── Вспомогательные методы ────────────────────────────────────────────────
    def _set_status(self, text: str, kind: str = "info") -> None:
        """Обновляет статусную строку (можно вызывать из любого потока)."""
        colors = {
            "ok":      self.OK_FG,
            "error":   self.ERR_FG,
            "warning": self.WARN_FG,
            "info":    self.HINT_FG,
        }

        def _update() -> None:
            self._status_var.set(text)
            self._status_lbl.configure(fg=colors.get(kind, self.HINT_FG))

        self.after(0, _update)

    def _init_client(self) -> None:
        """Инициализирует GSheetClient по переменным окружения."""
        spreadsheet_id = os.getenv("GOOGLE_SPREADSHEET_ID", "").strip()
        sa_raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip() or None

        if sa_raw:
            p = Path(sa_raw).expanduser()
            if not p.is_absolute():
                candidate = Path(__file__).parent / sa_raw
                if candidate.exists():
                    sa_raw = str(candidate)
                elif p.exists():
                    sa_raw = str(p)

        if spreadsheet_id and sa_raw:
            try:
                self._client = GSheetClient(
                    spreadsheet_id=spreadsheet_id,
                    service_account_file=sa_raw,
                )
                self._set_status("✅ Подключено к Google Sheets", "ok")
            except Exception as exc:
                self._set_status(f"❌ Ошибка подключения: {exc}", "error")
        else:
            self._set_status(
                "⚠️  Настройте GOOGLE_SPREADSHEET_ID и "
                "GOOGLE_SERVICE_ACCOUNT_FILE в файле .env",
                "warning",
            )

    def _get_dates(self) -> Tuple[date, date]:
        """Считывает даты из виджетов."""
        if _HAS_CAL:
            return self._date_from.get_date(), self._date_to.get_date()
        try:
            d_from = datetime.strptime(self._date_from_var.get().strip(), "%d.%m.%Y").date()
            d_to   = datetime.strptime(self._date_to_var.get().strip(), "%d.%m.%Y").date()
            return d_from, d_to
        except ValueError as exc:
            raise ValueError(
                "Неверный формат даты. Используйте дд.мм.гггг"
            ) from exc

    # ── Обработчик кнопки ─────────────────────────────────────────────────────
    def _on_generate(self) -> None:
        if self._busy:
            return
        if self._client is None:
            messagebox.showerror(
                "Нет подключения",
                "Подключение к Google Sheets не установлено.\n"
                "Проверьте файл .env и перезапустите приложение.",
            )
            return

        # Валидация
        company = self._company_var.get().strip()
        department = self._dept_var.get().strip()
        responsible = self._resp_var.get().strip()

        if not company:
            messagebox.showwarning("Заполните поле", "Введите название организации.")
            return
        if not responsible:
            messagebox.showwarning("Заполните поле", "Введите ФИО ответственного.")
            return

        try:
            date_from, date_to = self._get_dates()
        except ValueError as exc:
            messagebox.showerror("Неверная дата", str(exc))
            return

        if date_from > date_to:
            messagebox.showwarning("Неверный период", "Дата «с» должна быть не позже даты «по».")
            return

        # Название листа
        user_name = self._sheet_name_var.get().strip()
        if user_name:
            sheet_title = user_name[:100]
        else:
            dept_short = department[:12].rstrip()
            sheet_title = f"Отчёт — {dept_short} — {datetime.now().strftime('%d.%m.%y %H-%M')}"

        # Запускаем в фоновом потоке, чтобы не замораживать UI
        self._busy = True
        self._btn.configure(state="disabled", text="⏳  Формируем отчёт…")
        self._set_status("⏳ Создаём лист и записываем данные…", "info")

        def _worker() -> None:
            try:
                title = build_report_in_sheet(
                    client=self._client,          # type: ignore[arg-type]
                    company=company,
                    department=department,
                    responsible=responsible,
                    date_from=date_from,
                    date_to=date_to,
                    sheet_title=sheet_title,
                )
                self._set_status(
                    f"✅ Отчёт успешно создан на листе «{title}»", "ok"
                )
                self.after(0, lambda: messagebox.showinfo(
                    "Готово",
                    f"Отчёт записан на лист:\n«{title}»\n\n"
                    "Откройте Google Таблицу, чтобы посмотреть результат.",
                ))
            except Exception as exc:  # noqa: BLE001
                self._set_status(f"❌ Ошибка: {exc}", "error")
                self.after(0, lambda: messagebox.showerror("Ошибка", str(exc)))
            finally:
                self._busy = False
                self.after(0, lambda: self._btn.configure(
                    state="normal", text="🚀  Сформировать отчёт"
                ))

        threading.Thread(target=_worker, daemon=True).start()


# ══════════════════════════════════════════════════════════════════════════════
#  Точка входа
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = ReportApp()
    app.mainloop()
