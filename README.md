# GSheetAPI

Проект объединяет две части:

- `Mini CRM` на `FastAPI + Tkinter + SQLite`
- интеграция с `Google Drive` и `Google Sheets` для выгрузки отчётов из GUI

Сейчас основная пользовательская история такая: CRM запускается локально, в GUI есть таблицы клиентов, сделок и задач, а каждую таблицу можно выгрузить в отдельный красиво оформленный Google Sheet от имени личного Google-аккаунта пользователя через OAuth.

## Структура проекта

```text
GSheetAPI/
├── .dockerignore                    # исключения для docker build context
├── crm.py                         # основной запуск CRM API из корня
├── crm.db                         # SQLite база CRM
├── docker-compose.yml             # запуск FastAPI backend через Docker Compose
├── seed_crm_data.py               # заполнение CRM тестовыми данными через API
├── README.md
├── requirements.txt
├── token.json                     # OAuth token cache (создаётся автоматически)
├── crm/
│   ├── Dockerfile                 # образ FastAPI backend для compose
│   ├── crm.py                     # альтернативная точка входа API внутри папки crm/
│   ├── crm_api.py                 # FastAPI API
│   ├── crm_db.py                  # SQLite data layer / CRUD / поиск / статистика
│   ├── crm_models.py              # Pydantic-модели и SQL-схемы
│   ├── crm_ui.py                  # Tkinter GUI Mini CRM + экспорт в Google
│   └── google_settings.json       # сохранённые настройки Google-интеграции GUI
├── drive/
│   ├── gdrive_api.py              # Google Drive через сервисный аккаунт
│   └── gdrive_oauth_api.py        # Google Drive через OAuth 2.0
├── sheets/
│   ├── gsheet_api.py              # Google Sheets через сервисный аккаунт
│   └── report_app.py              # отдельное demo-приложение отчётов
├── google_integration/
│   ├── google_drive.py            # GUI-обвязка для Drive OAuth
│   └── google_sheets.py           # GUI-обвязка для записи и форматирования отчётов
└── keys/
    ├── client_secret_*.json
    └── service-account.json
```

## Что умеет проект

### Mini CRM

- вести клиентов, сделки и задачи
- фильтровать и искать записи в GUI
- смотреть сводную статистику на дашборде
- создавать, редактировать, удалять и архивировать записи
- работать локально без внешнего сервера

### Google-интеграция

- открывать окно настроек Google прямо из GUI
- сохранять пути к `JSON`-файлам и `folder_id` в `crm/google_settings.json`
- авторизовывать пользователя через OAuth в браузере
- создавать отдельный Google Sheet для каждой таблицы CRM
- записывать данные в новый Sheet и применять форматирование
- показывать ссылку на созданный отчёт и открывать его в браузере

### Тестовые данные

- скрипт `seed_crm_data.py` очищает текущие данные CRM
- затем заново создаёт около `30` клиентов, `30` сделок и `30` задач
- данные похожи на реальные: ФИО, компании, телефоны, email, названия сделок и задач
- сделки создаются только в валюте `RUB`

## Быстрый старт

### 1. Создать и активировать виртуальное окружение

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
```

### 2. Установить зависимости

```powershell
pip install -r requirements.txt
```

### 3. Подготовить Google-ключи

Положите в `keys/`:

- `client_secret_*.json` для OAuth Desktop App
- JSON сервисного аккаунта для модулей `drive/` и `sheets/`

### 4. Запустить CRM API

Запускать нужно из корня проекта:

```powershell
venv\Scripts\python.exe crm.py
```

После запуска доступны:

- `http://127.0.0.1:8000/docs`
- `http://127.0.0.1:8000/redoc`
- `http://127.0.0.1:8000/dashboard`

### 5. Запустить GUI

В отдельном окне:

```powershell
cd crm
..\venv\Scripts\python.exe crm_ui.py
```

### Альтернатива: запуск API через Docker Compose

Если не хотите поднимать локальное `venv`, можно запустить только backend через `Docker Compose`:

```powershell
docker compose up --build
```

Или в фоне:

```powershell
docker compose up --build -d
```

После запуска доступны:

- `http://127.0.0.1:8000/docs`
- `http://127.0.0.1:8000/redoc`
- `http://127.0.0.1:8000/dashboard`

Что делает `docker-compose.yml`:

- собирает образ из `crm/Dockerfile`
- поднимает сервис `crm_api`
- пробрасывает порт `8000:8000`
- хранит SQLite-базу в именованном volume `crm_db_data`
- монтирует `./crm` в контейнер, поэтому изменения в `*.py` подхватываются через hot reload

Остановка:

```powershell
docker compose down
```

С удалением volume базы:

```powershell
docker compose down -v
```

> `Docker Compose` в этом проекте поднимает только FastAPI API. `Tkinter GUI` по-прежнему запускается локально на хосте.

## Основные команды

```powershell
# CRM API
venv\Scripts\python.exe crm.py

# CRM API с reload
venv\Scripts\python.exe crm.py --reload

# Tkinter GUI CRM
cd crm
..\venv\Scripts\python.exe crm_ui.py

# CRM API через Docker Compose
docker compose up --build

# Остановить Docker Compose
docker compose down

# Заполнить CRM тестовыми данными через API
cd ..
venv\Scripts\python.exe seed_crm_data.py

# Проверить сервисный Google Sheets client
venv\Scripts\python.exe -X utf8 sheets\gsheet_api.py

# Проверить сервисный Google Drive client
venv\Scripts\python.exe -X utf8 drive\gdrive_api.py

# Проверить OAuth Drive client
venv\Scripts\python.exe -X utf8 drive\gdrive_oauth_api.py

# Отдельное demo-приложение отчётов
venv\Scripts\python.exe sheets\report_app.py
```

> На Windows флаг `-X utf8` полезен для корректной кириллицы в консоли.

## CRM API

Основной backend находится в `crm/crm_api.py`.

### Сущности

- `Clients`
- `Deals`
- `Tasks`

### Основные endpoints

- `GET /dashboard`
- `GET /clients`
- `POST /clients`
- `PUT /clients/{id}`
- `PATCH /clients/{id}/archive`
- `GET /deals`
- `POST /deals`
- `PUT /deals/{id}`
- `PATCH /deals/{id}/status`
- `GET /tasks`
- `POST /tasks`
- `PUT /tasks/{id}`
- `PATCH /tasks/{id}/complete`

### Поиск

Поиск в CRM работает по видимым полям таблиц GUI:

- `Клиенты`: `name`, `email`, `company`, `phone`
- `Сделки`: `title`
- `Задачи`: `title`

Для кириллицы поиск сделан через `Python casefold()`, а не через SQLite `LIKE`, чтобы корректно находить, например:

- `попов` -> `Попов`
- `внедрение` -> `Внедрение`
- `согласовать` -> `Согласовать`

## Экспорт отчётов в Google

Экспорт реализован в `crm/crm_ui.py` через модули:

- `google_integration/google_drive.py`
- `google_integration/google_sheets.py`

### Как это работает

1. Пользователь нажимает кнопку выгрузки на вкладке с таблицей.
2. GUI собирает массив данных из текущей таблицы.
3. Через OAuth создаётся новый Google Sheet в личном Google Drive пользователя.
4. В новый Sheet записываются данные и сводка.
5. Применяется форматирование: заголовки, секции, полосы строк, границы, заморозка шапки.
6. GUI показывает ссылку на созданный отчёт.

### Где настраивается интеграция

В самом GUI:

- кнопка `Настройки Google`
- путь к `service account JSON`
- путь к `client_secret JSON`
- `folder_id`
- кнопка `Вставить` рядом с `folder_id`

Настройки сохраняются в `crm/google_settings.json`.

Дополнительно сохраняются:

- ID последнего отчёта
- название последнего отчёта
- ссылка на последний отчёт

## Модули Google API

### `sheets/gsheet_api.py`

Клиент Google Sheets через сервисный аккаунт.

Ключевые методы:

- `read_values()`
- `append_rows()`
- `update_range()`
- `batch_update_ranges()`
- `clear_range()`
- `delete_rows()`
- `apply_requests()`
- `create_sheet()`
- `delete_sheet()`

### `drive/gdrive_api.py`

Клиент Google Drive через сервисный аккаунт.

Ключевые методы:

- `list_files()`
- `get_file_metadata()`
- `download_file()`
- `export_google_doc()`
- `create_folder()`
- `upload_file()`
- `create_google_sheet()`
- `create_google_doc()`
- `rename_file()`
- `move_file()`
- `update_file_content()`
- `update_file_metadata()`
- `delete_file()`
- `trash_file()`

### `drive/gdrive_oauth_api.py`

Google Drive от имени личного аккаунта пользователя через OAuth 2.0.

Используется, когда файл должен появляться в личном Drive пользователя, а не в хранилище сервисного аккаунта.

### `google_integration/google_drive.py`

GUI-обвязка над OAuth-клиентом:

- получает OAuth credentials
- кэширует токен в `token.json`
- создаёт Google Sheets из GUI

### `google_integration/google_sheets.py`

GUI-обвязка над `GSheetClient`:

- пишет табличный отчёт в новый Sheet
- добавляет сводку
- форматирует документ

## Скрипт `seed_crm_data.py`

Скрипт полезен для быстрого наполнения демо-данными.

Что делает:

1. удаляет все текущие `tasks`
2. удаляет все текущие `deals`
3. удаляет всех текущих `clients`
4. создаёт новые тестовые записи через HTTP API

Запуск:

```powershell
venv\Scripts\python.exe seed_crm_data.py
```

Ожидается, что API уже работает на `http://127.0.0.1:8000`.

## Зависимости

Основные пакеты:

- `fastapi`
- `uvicorn`
- `requests`
- `google-api-python-client`
- `google-auth`
- `google-auth-httplib2`
- `google-auth-oauthlib`
- `python-dotenv`
- `tkcalendar`

## Безопасность

Не коммитьте в git:

- `token.json`
- `keys/*.json`
- `.env`
- локальные файлы БД, если они содержат реальные данные

Рекомендуемый `.gitignore`:

```gitignore
.env
token.json
keys/*.json
venv/
__pycache__/
*.pyc
```

## Требования

- Python `3.11+`
- Windows / macOS / Linux
- доступ в интернет для Google API
- Google Cloud project с включёнными `Google Sheets API` и `Google Drive API`
- OAuth Desktop App для GUI-экспорта
- сервисный аккаунт для низкоуровневых модулей `drive/` и `sheets/`
