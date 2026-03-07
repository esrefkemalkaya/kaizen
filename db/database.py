import sqlite3
from pathlib import Path

DB_PATH = Path.home() / ".kaizen_drilling" / "kaizen.db"


def get_connection() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_connection()
    c = conn.cursor()

    c.executescript("""
        CREATE TABLE IF NOT EXISTS projects (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            name        TEXT NOT NULL,
            location    TEXT,
            created_at  TEXT DEFAULT (date('now'))
        );

        CREATE TABLE IF NOT EXISTS contractors (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id        INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
            name              TEXT NOT NULL,
            type              TEXT NOT NULL CHECK(type IN ('underground', 'surface')),
            rate_per_meter    REAL NOT NULL DEFAULT 0,
            standby_hour_rate REAL NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS drilling_entries (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            contractor_id  INTEGER NOT NULL REFERENCES contractors(id) ON DELETE CASCADE,
            month          INTEGER NOT NULL,
            year           INTEGER NOT NULL,
            hole_id        TEXT NOT NULL,
            meters_drilled REAL NOT NULL DEFAULT 0,
            standby_hours  REAL NOT NULL DEFAULT 0,
            start_date     TEXT NOT NULL DEFAULT '',
            end_date       TEXT NOT NULL DEFAULT '',
            start_depth    REAL NOT NULL DEFAULT 0,
            end_depth      REAL NOT NULL DEFAULT 0,
            rig_name       TEXT NOT NULL DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS ppe_charges (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            contractor_id   INTEGER NOT NULL REFERENCES contractors(id) ON DELETE CASCADE,
            month           INTEGER NOT NULL,
            year            INTEGER NOT NULL,
            material_code   TEXT NOT NULL DEFAULT '',
            item_name       TEXT NOT NULL,
            quantity        REAL NOT NULL DEFAULT 0,
            unit_of_measure TEXT NOT NULL DEFAULT '',
            entry_date      TEXT NOT NULL DEFAULT '',
            unit_price      REAL NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS diesel_charges (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            contractor_id INTEGER NOT NULL REFERENCES contractors(id) ON DELETE CASCADE,
            month         INTEGER NOT NULL,
            year          INTEGER NOT NULL,
            description   TEXT NOT NULL,
            quantity      REAL NOT NULL DEFAULT 0,
            unit_price    REAL NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS standby_entries (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            contractor_id INTEGER NOT NULL REFERENCES contractors(id) ON DELETE CASCADE,
            month         INTEGER NOT NULL,
            year          INTEGER NOT NULL,
            entry_date    TEXT NOT NULL DEFAULT '',
            hole_id       TEXT NOT NULL DEFAULT '',
            start_time    TEXT NOT NULL DEFAULT '',
            end_time      TEXT NOT NULL DEFAULT '',
            standby_type  TEXT NOT NULL DEFAULT '',
            description   TEXT NOT NULL DEFAULT '',
            hours         REAL NOT NULL DEFAULT 0,
            rig_name      TEXT NOT NULL DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS settings (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS period_settings (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            contractor_id   INTEGER NOT NULL REFERENCES contractors(id) ON DELETE CASCADE,
            month           INTEGER NOT NULL,
            year            INTEGER NOT NULL,
            donem_adi       TEXT    NOT NULL DEFAULT '',
            exchange_rate   REAL    NOT NULL DEFAULT 0,
            kuyuda_kalan    REAL    NOT NULL DEFAULT 0,
            target_geo1     REAL    NOT NULL DEFAULT 700,
            target_geo2     REAL    NOT NULL DEFAULT 700,
            target_geo3     REAL    NOT NULL DEFAULT 700,
            target_geo5     REAL    NOT NULL DEFAULT 700,
            standby_rate    REAL    NOT NULL DEFAULT 75,
            UNIQUE(contractor_id, month, year)
        );
    """)

    conn.commit()
    _migrate_db(conn)
    conn.close()


def _migrate_db(conn):
    """Add any missing columns to existing databases."""
    c = conn.cursor()

    # drilling_entries new columns
    existing = {r[1] for r in c.execute("PRAGMA table_info(drilling_entries)").fetchall()}
    for col, defn in [
        ("start_date",  "TEXT NOT NULL DEFAULT ''"),
        ("end_date",    "TEXT NOT NULL DEFAULT ''"),
        ("start_depth", "REAL NOT NULL DEFAULT 0"),
        ("end_depth",   "REAL NOT NULL DEFAULT 0"),
        ("rig_name",    "TEXT NOT NULL DEFAULT ''"),
    ]:
        if col not in existing:
            c.execute(f"ALTER TABLE drilling_entries ADD COLUMN {col} {defn}")

    # ppe_charges new columns
    existing = {r[1] for r in c.execute("PRAGMA table_info(ppe_charges)").fetchall()}
    for col, defn in [
        ("material_code",   "TEXT NOT NULL DEFAULT ''"),
        ("unit_of_measure", "TEXT NOT NULL DEFAULT ''"),
        ("entry_date",      "TEXT NOT NULL DEFAULT ''"),
    ]:
        if col not in existing:
            c.execute(f"ALTER TABLE ppe_charges ADD COLUMN {col} {defn}")

    # standby_entries new columns
    existing = {r[1] for r in c.execute("PRAGMA table_info(standby_entries)").fetchall()}
    for col, defn in [
        ("entry_date",   "TEXT NOT NULL DEFAULT ''"),
        ("hole_id",      "TEXT NOT NULL DEFAULT ''"),
        ("start_time",   "TEXT NOT NULL DEFAULT ''"),
        ("end_time",     "TEXT NOT NULL DEFAULT ''"),
        ("standby_type", "TEXT NOT NULL DEFAULT ''"),
        ("rig_name",     "TEXT NOT NULL DEFAULT ''"),
    ]:
        if col not in existing:
            c.execute(f"ALTER TABLE standby_entries ADD COLUMN {col} {defn}")

    conn.commit()
