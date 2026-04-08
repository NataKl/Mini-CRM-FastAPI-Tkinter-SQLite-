"""
crm.py — точка входа мини-CRM
──────────────────────────────
Запускает FastAPI-сервер через uvicorn.

Использование:
    venv\\Scripts\\python.exe crm.py
    venv\\Scripts\\python.exe crm.py --host 0.0.0.0 --port 8080

После запуска:
    Swagger UI  → http://127.0.0.1:8000/docs
    ReDoc       → http://127.0.0.1:8000/redoc
    Dashboard   → http://127.0.0.1:8000/dashboard
"""

import argparse
import uvicorn


def main() -> None:
    parser = argparse.ArgumentParser(description="Mini CRM — FastAPI server")
    parser.add_argument("--host",   default="127.0.0.1", help="Bind host (default: 127.0.0.1)")
    parser.add_argument("--port",   default=8000, type=int, help="Bind port (default: 8000)")
    parser.add_argument("--reload", action="store_true",   help="Enable auto-reload on code changes")
    parser.add_argument("--db",     default="crm.db",      help="Path to SQLite DB file (default: crm.db)")
    args = parser.parse_args()

    import os
    os.environ.setdefault("CRM_DB_PATH", args.db)

    print(f"[CRM]  API:        http://{args.host}:{args.port}")
    print(f"[CRM]  Swagger UI: http://{args.host}:{args.port}/docs")
    print(f"[CRM]  DB:         {args.db}")

    uvicorn.run(
        "crm_api:app",
        host=args.host,
        port=args.port,
        reload=args.reload,
    )


if __name__ == "__main__":
    main()
