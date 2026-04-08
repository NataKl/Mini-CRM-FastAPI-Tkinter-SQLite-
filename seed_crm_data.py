"""Fill the CRM API with realistic-looking random demo data."""

from __future__ import annotations

import random
import time
from datetime import date, timedelta
from typing import Dict, List, Optional

import requests


API_URL = "http://127.0.0.1:8000"
CLIENTS_TO_CREATE = 30
DEALS_TO_CREATE = 30
TASKS_TO_CREATE = 30

FIRST_NAMES = [
    "Александр", "Алексей", "Андрей", "Анна", "Виктор", "Виктория", "Дмитрий",
    "Евгений", "Екатерина", "Елена", "Иван", "Ирина", "Кирилл", "Мария",
    "Максим", "Наталья", "Ольга", "Павел", "Светлана", "Сергей", "Татьяна",
]
LAST_NAMES = [
    "Иванов", "Петров", "Сидоров", "Смирнов", "Кузнецов", "Попов", "Васильев",
    "Морозов", "Новиков", "Федоров", "Волков", "Лебедев", "Козлова", "Соколова",
    "Орлова", "Николаева", "Зайцева", "Павлова", "Семёнова", "Громова",
]
PATRONYMICS = [
    "Александрович", "Алексеевич", "Андреевич", "Викторович", "Дмитриевич",
    "Евгеньевич", "Иванович", "Кириллович", "Максимович", "Павлович",
    "Сергеевич", "Александровна", "Алексеевна", "Андреевна", "Викторовна",
    "Дмитриевна", "Евгеньевна", "Ивановна", "Кирилловна", "Павловна", "Сергеевна",
]
COMPANY_PREFIXES = [
    "Техно", "Север", "Вектор", "Альфа", "Профи", "Бизнес", "Гранд", "Пром",
    "Лидер", "Нева", "Строй", "Фин", "Логистик", "Оптима", "Смарт", "Интегра",
]
COMPANY_SUFFIXES = [
    "Софт", "Сервис", "Групп", "Снаб", "Партнёр", "Лайн", "Системы", "Решения",
    "Трейд", "Логистика", "Консалт", "Плюс", "Маркет", "Энерго", "Финанс",
]
DOMAINS = [
    "mail.ru", "yandex.ru", "company.ru", "inbox.ru", "corp.pro", "business.io",
]
CLIENT_SOURCES = [
    "Сайт", "Рекомендация", "Холодный звонок", "Telegram", "Выставка",
    "Контекстная реклама", "Повторное обращение", "Партнёрский канал",
]
CLIENT_STATUSES = ["lead", "lead", "active", "active", "active", "vip", "inactive"]
DEAL_STATUSES = ["new", "new", "in_progress", "in_progress", "won", "lost", "on_hold"]
TASK_STATUSES = ["pending", "pending", "pending", "done", "cancelled"]
TASK_PRIORITIES = ["low", "medium", "medium", "high"]
DEAL_PRODUCTS = [
    "внедрение CRM", "поставка лицензий", "техническая поддержка", "интеграция телефонии",
    "аудит бизнес-процессов", "разработка дашбордов", "автоматизация отдела продаж",
    "обслуживание серверов", "внедрение BI", "проект цифровизации документооборота",
]
TASK_TITLES = [
    "Подготовить коммерческое предложение",
    "Согласовать договор",
    "Провести демонстрацию продукта",
    "Позвонить клиенту и уточнить требования",
    "Отправить счёт на оплату",
    "Подготовить презентацию для встречи",
    "Проверить статус согласования",
    "Запросить обратную связь после пилота",
    "Обновить данные в CRM",
    "Согласовать дату следующего созвона",
]
TASK_DESCRIPTIONS = [
    "Нужно согласовать детали и зафиксировать итог в CRM.",
    "Уточнить потребности клиента и следующий шаг по воронке.",
    "Подготовить материалы и отправить подтверждение на почту.",
    "Проверить дедлайны и предупредить ответственного менеджера.",
    "Собрать вводные от клиента и передать в проектную команду.",
]
NOTES = [
    "Интересуется быстрым запуском и поддержкой.",
    "Решение выбирается между двумя подрядчиками.",
    "Есть внутренний бюджет и согласование у директора.",
    "Клиент просил коммерческое предложение в расширенном варианте.",
    "Контакт тёплый, ранее уже обсуждали сотрудничество.",
]


class ApiClient:
    def __init__(self, base_url: str = API_URL) -> None:
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers["Content-Type"] = "application/json"

    def _url(self, path: str) -> str:
        return f"{self.base_url}/{path.lstrip('/')}"

    def get(self, path: str, **params):
        response = self.session.get(self._url(path), params=params or None, timeout=10)
        response.raise_for_status()
        return response.json()

    def post(self, path: str, body: Dict):
        response = self.session.post(self._url(path), json=body, timeout=10)
        response.raise_for_status()
        return response.json()

    def delete(self, path: str) -> None:
        response = self.session.delete(self._url(path), timeout=10)
        response.raise_for_status()


def _slug(text: str) -> str:
    repl = {
        "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
        "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
        "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
        "ф": "f", "х": "h", "ц": "c", "ч": "ch", "ш": "sh", "щ": "sch", "ъ": "",
        "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya", " ": ".", "-": ".",
    }
    return "".join(repl.get(ch, ch) for ch in text.lower() if ch.isalnum() or ch in " -").strip(".")


def _random_company() -> str:
    legal = random.choice(["ООО", "АО", "ИП"])
    return f"{legal} «{random.choice(COMPANY_PREFIXES)}{random.choice(COMPANY_SUFFIXES)}»"


def _random_full_name() -> str:
    return f"{random.choice(LAST_NAMES)} {random.choice(FIRST_NAMES)} {random.choice(PATRONYMICS)}"


def _random_phone() -> str:
    # Не используем ведущий "+" — при экспорте в Google Sheets это может
    # интерпретироваться как формула и давать #ERROR!.
    return f"8 ({random.randint(900, 999)}) {random.randint(100, 999)}-{random.randint(10, 99)}-{random.randint(10, 99)}"


def _random_email(full_name: str, index: int) -> str:
    parts = full_name.split()
    local = ".".join(_slug(part) for part in parts[:2] if part)
    return f"{local}.{index}@{random.choice(DOMAINS)}"


def _random_client_payload(index: int) -> Dict:
    full_name = _random_full_name()
    return {
        "name": full_name,
        "email": _random_email(full_name, index),
        "phone": _random_phone(),
        "company": _random_company(),
        "source": random.choice(CLIENT_SOURCES),
        "status": random.choice(CLIENT_STATUSES),
        "notes": random.choice(NOTES),
    }


def _random_due_date(days_from: int, days_to: int) -> str:
    return (date.today() + timedelta(days=random.randint(days_from, days_to))).isoformat()


def _random_deal_payload(client: Dict, index: int) -> Dict:
    product = random.choice(DEAL_PRODUCTS)
    amount = random.randint(80_000, 4_500_000)
    return {
        "title": f"{product.capitalize()} для {client['company']}",
        "client_id": client["id"],
        "amount": float(amount),
        "currency": "RUB",
        "status": random.choice(DEAL_STATUSES),
        "description": (
            f"Сделка #{index + 1}: {product}. "
            f"Контактное лицо: {client['name']}. "
            f"{random.choice(NOTES)}"
        ),
        "due_date": _random_due_date(-20, 60),
    }


def _random_task_payload(client: Dict, deal: Optional[Dict], index: int) -> Dict:
    title = random.choice(TASK_TITLES)
    return {
        "title": f"{title} #{index + 1}",
        "description": random.choice(TASK_DESCRIPTIONS),
        "client_id": client["id"],
        "deal_id": deal["id"] if deal else None,
        "status": random.choice(TASK_STATUSES),
        "priority": random.choice(TASK_PRIORITIES),
        "due_date": _random_due_date(-10, 30),
    }


def _clear_existing_data(api: ApiClient) -> None:
    print("Очищаю существующие данные CRM...")

    tasks = api.get("/tasks", limit=1000)
    for task in tasks:
        api.delete(f"/tasks/{task['id']}")
    print(f"  Удалено задач:   {len(tasks)}")

    deals = api.get("/deals", limit=1000)
    for deal in deals:
        api.delete(f"/deals/{deal['id']}")
    print(f"  Удалено сделок:  {len(deals)}")

    clients = api.get("/clients", limit=1000)
    for client in clients:
        api.delete(f"/clients/{client['id']}")
    print(f"  Удалено клиентов: {len(clients)}")


def main() -> None:
    random.seed()
    api = ApiClient()

    try:
        api.get("/dashboard")
    except Exception as exc:  # noqa: BLE001
        raise SystemExit(
            "CRM API недоступно. Запустите сервер на http://127.0.0.1:8000 и повторите.\n"
            f"Подробности: {exc}"
        ) from exc

    _clear_existing_data(api)

    created_clients: List[Dict] = []
    created_deals: List[Dict] = []
    created_tasks: List[Dict] = []
    nonce = int(time.time())

    print(f"Создаю клиентов: {CLIENTS_TO_CREATE}")
    for idx in range(CLIENTS_TO_CREATE):
        payload = _random_client_payload(nonce + idx)
        client = api.post("/clients", payload)
        created_clients.append(client)
        print(f"  + client #{client['id']}: {client['name']} / {client.get('company')}")

    print(f"\nСоздаю сделки: {DEALS_TO_CREATE}")
    for idx in range(DEALS_TO_CREATE):
        client = random.choice(created_clients)
        payload = _random_deal_payload(client, idx)
        deal = api.post("/deals", payload)
        created_deals.append(deal)
        print(f"  + deal #{deal['id']}: {deal['title']}")

    print(f"\nСоздаю задачи: {TASKS_TO_CREATE}")
    for idx in range(TASKS_TO_CREATE):
        client = random.choice(created_clients)
        related_deals = [deal for deal in created_deals if deal.get("client_id") == client["id"]]
        deal = random.choice(related_deals) if related_deals and random.random() < 0.8 else None
        payload = _random_task_payload(client, deal, idx)
        task = api.post("/tasks", payload)
        created_tasks.append(task)
        print(f"  + task #{task['id']}: {task['title']}")

    print("\nГотово.")
    print(f"Создано клиентов: {len(created_clients)}")
    print(f"Создано сделок:   {len(created_deals)}")
    print(f"Создано задач:    {len(created_tasks)}")


if __name__ == "__main__":
    main()
