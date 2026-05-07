# AGENTS.md

## Project
Single-file Flask app (`app.py`) for tracking laboratory equipment repairs. SQLite database (`laboratorio.db`). Python 3.13 via `venv/`.

## Run
```
venv\Scripts\activate
flask run
```
No `requirements.txt` exists. Installed packages: `flask`, `openpyxl`, `werkzeug`. Install with `pip install flask openpyxl` if venv is missing deps.

## Database
- DB file: `laboratorio.db` (SQLite, WAL mode, foreign keys ON)
- Schema is created at runtime by `ensure_support_tables()` in `app.py` — there is no separate migration system
- `schema.sql` exists but is **out of date** — trust the `CREATE TABLE` statements inside `app.py` instead
- Custom SQL function `sortable_date()` registered per-connection for sorting inconsistent D/M/YYYY dates

## Auth
- Default admin: `admin` / `Admin@Lab2024!` (created on first `init_db()` call)
- Two roles: `admin`, `tecnico`
- Single-session enforcement: a second login forces out the first session (unless `force=true` is posted)
- Session tracked via `session_id` + `last_ip` columns in `Usuarios` table

## Key conventions
- Dates stored as `D/M/YYYY` or `D/M/YYYY H:MM` strings (Argentine locale), **not** ISO format
- Prices in `EquiposCatalogo.valor` stored as plain numeric strings (no `$` prefix); displayed with `format_price` filter using `.` as thousands separator and no decimals
- Catalog values are auto-created on insert via `ensure_catalog_value()` — no validation step needed
- `app (copia).py` is a backup copy — do not edit it

## CSV seed files
- `reparaciones.csv` (~10k rows), `equipos.csv` (~495 rows), `familias.csv` (~52 rows)
- No import script exists in the repo — these are reference data, not loaded by the app

## Structure
```
app.py                  # Single entrypoint — all routes, DB logic, and helpers
templates/              # 9 Jinja2 templates (index, login, dashboard, configuracion, etc.)
laboratorio.db          # SQLite database (generated at runtime)
schema.sql              # OUTDATED — ignore, use app.py schema
venv/                   # Python 3.13 virtualenv
*.csv                   # Reference/seed data (not auto-loaded)
requirements.txt       # Dependencies (Flask, Flask-WTF, openpyxl, werkzeug)
backups/               # Automatic backups directory (created at runtime)
```

## No test/lint/typecheck setup
This repo has no tests, linter, formatter, or type checker. Do not invent them unless asked.

## Security Features (Implemented)
- CSRF Protection via Flask-WTF
- Rate Limiting: 5 login attempts per 5 minutes, then 15 min block
- Secure cookies: HTTPONLY, SAMESITE=Lax
- Audit logging: All actions registered in `Historial` table

## Historial de Acciones
- Route: `/configuracion/historial`
- Retains records for 1 year
- Records: LOGIN, LOGOUT, CREAR, EDITAR, ELIMINAR, CAMBIO_ESTADO, ASIGNAR_TECNICO, BACKUP, CONFIGURACION

## Backup
- Automatic: Every 24hs (thread running in background)
- Manual: `/backup/crear` and `/backup/descargar` routes
- Location: `backups/laboratorio_YYYYMMDD_HHMMSS.db`

## Health Check
- Route: `/health`
- Returns JSON with status and timestamp
