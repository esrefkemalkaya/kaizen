import sqlite3
import os
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
            standby_hours  REAL NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS ppe_charges (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            contractor_id INTEGER NOT NULL REFERENCES contractors(id) ON DELETE CASCADE,
            month         INTEGER NOT NULL,
            year          INTEGER NOT NULL,
            item_name     TEXT NOT NULL,
            quantity      REAL NOT NULL DEFAULT 0,
            unit_price    REAL NOT NULL DEFAULT 0
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
    """)

    conn.commit()
    conn.close()
