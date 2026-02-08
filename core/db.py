# -*- coding: utf-8 -*-
"""Database abstraction layer — supports both SQLite (dev) and PostgreSQL (production)."""
from __future__ import annotations

import os
import sqlite3
from typing import Any, Optional, Sequence, Union

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
DATABASE_URL: str = os.environ.get("DATABASE_URL", "")
_USE_POSTGRES: bool = DATABASE_URL.startswith("postgresql://") or DATABASE_URL.startswith("postgres://")

DB_FILENAME = "forms.db"

if _USE_POSTGRES:
    import psycopg2
    import psycopg2.extras


def is_postgres() -> bool:
    return _USE_POSTGRES


# ---------------------------------------------------------------------------
# Unified cursor / connection wrappers
# ---------------------------------------------------------------------------

class Cursor:
    """Thin wrapper that exposes *lastrowid*, *rowcount*, fetchone, fetchall."""

    def __init__(self, cursor: Any, lastrowid: Optional[int] = None):
        self._cursor = cursor
        self.lastrowid: Optional[int] = lastrowid
        self.rowcount: int = cursor.rowcount

    def fetchone(self):
        return self._cursor.fetchone()

    def fetchall(self):
        return self._cursor.fetchall()


class Connection:
    """Unified database connection that works with both SQLite and PostgreSQL."""

    def __init__(self, conn: Any, *, postgres: bool = False):
        self._conn = conn
        self._postgres = postgres

    # -- query helpers -----------------------------------------------------

    def execute(self, query: str, params: Union[tuple, Sequence] = ()) -> Cursor:
        if self._postgres:
            query = _convert_placeholders(query)
            cur = self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(query, params or None)
            return Cursor(cur)
        else:
            cur = self._conn.execute(query, params)
            return Cursor(cur, lastrowid=cur.lastrowid)

    def execute_returning_id(
        self, query: str, params: Union[tuple, Sequence] = ()
    ) -> Optional[int]:
        """Execute an INSERT statement and return the new row's *id*."""
        if self._postgres:
            query = _convert_placeholders(query).rstrip().rstrip(";")
            if "RETURNING" not in query.upper():
                query += " RETURNING id"
            cur = self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(query, params or None)
            row = cur.fetchone()
            return row["id"] if row else None
        else:
            cur = self._conn.execute(query, params)
            return cur.lastrowid

    # -- transaction helpers -----------------------------------------------

    def commit(self) -> None:
        self._conn.commit()

    # -- context manager ---------------------------------------------------

    def __enter__(self) -> "Connection":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._postgres:
            if exc_type:
                self._conn.rollback()
            self._conn.close()
        else:
            # sqlite3 context manager: commit on success, rollback on error
            self._conn.__exit__(exc_type, exc_val, exc_tb)
        return False


# ---------------------------------------------------------------------------
# Connection factory
# ---------------------------------------------------------------------------

_schema_initialized = False


def get_connection(base_path: str = ".") -> Connection:
    """Return a database connection (PostgreSQL or SQLite depending on config)."""
    global _schema_initialized

    if _USE_POSTGRES:
        conn = psycopg2.connect(DATABASE_URL)
        wrapped = Connection(conn, postgres=True)
    else:
        db_path = _sqlite_path(base_path)
        raw = sqlite3.connect(db_path)
        raw.row_factory = sqlite3.Row
        raw.execute("PRAGMA foreign_keys = ON")
        wrapped = Connection(raw, postgres=False)

    if not _schema_initialized:
        _ensure_schema(wrapped)
        wrapped.commit()
        _schema_initialized = True

    return wrapped


def reset_schema_flag() -> None:
    """Reset the schema-initialized flag (useful for tests)."""
    global _schema_initialized
    _schema_initialized = False


# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

def _ensure_schema(conn: Connection) -> None:
    if _USE_POSTGRES:
        _ensure_schema_postgres(conn)
    else:
        _ensure_schema_sqlite(conn)


def _ensure_schema_sqlite(conn: Connection) -> None:
    """Create tables using SQLite dialect."""

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name TEXT NOT NULL,
            email TEXT,
            phone TEXT,
            password_hash TEXT,
            role TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            portal_user_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS forms (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            form_no TEXT NOT NULL UNIQUE,
            tarih TEXT,
            tarih_iso TEXT,
            dok_no TEXT,
            rev_no TEXT,
            avans TEXT,
            taseron TEXT,
            gorev_tanimi TEXT,
            gorev_yeri TEXT,
            gorev_yeri_lower TEXT,
            gorev_il TEXT,
            gorev_ilce TEXT,
            gorev_firma TEXT,
            gorev_tarih TEXT,
            gorev_tarih_iso TEXT,
            yapilan_isler TEXT,
            gorev_ekleri TEXT,
            harcama_bildirimleri TEXT,
            yola_cikis_tarih TEXT,
            yola_cikis_tarih_iso TEXT,
            yola_cikis_saat TEXT,
            donus_tarih TEXT,
            donus_tarih_iso TEXT,
            donus_saat TEXT,
            calisma_baslangic_tarih TEXT,
            calisma_baslangic_tarih_iso TEXT,
            calisma_baslangic_saat TEXT,
            calisma_bitis_tarih TEXT,
            calisma_bitis_tarih_iso TEXT,
            calisma_bitis_saat TEXT,
            mola_suresi TEXT,
            arac_plaka TEXT,
            hazirlayan TEXT,
            durum TEXT,
            personel_1 TEXT,
            personel_2 TEXT,
            personel_3 TEXT,
            personel_4 TEXT,
            personel_5 TEXT,
            personel_search TEXT,
            last_step INTEGER DEFAULT 0,
            assigned_to_user_id INTEGER,
            assigned_by_user_id INTEGER,
            assigned_at TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(assigned_to_user_id) REFERENCES users(id) ON DELETE SET NULL,
            FOREIGN KEY(assigned_by_user_id) REFERENCES users(id) ON DELETE SET NULL
        )
        """
    )
    conn.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_forms_form_no ON forms (form_no)"
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS form_sequence (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            last_no INTEGER NOT NULL DEFAULT 0
        )
        """
    )
    conn.execute("INSERT OR IGNORE INTO form_sequence (id, last_no) VALUES (1, 0)")

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS task_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_name TEXT NOT NULL,
            customer_phone TEXT,
            customer_email TEXT,
            customer_address TEXT,
            request_description TEXT NOT NULL,
            requirements TEXT,
            urgency TEXT DEFAULT 'normal',
            requested_by_user_id INTEGER NOT NULL,
            status TEXT DEFAULT 'pending',
            notes TEXT,
            assigned_to_user_id INTEGER,
            converted_form_no TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            converted_at TEXT,
            FOREIGN KEY(requested_by_user_id) REFERENCES users(id),
            FOREIGN KEY(assigned_to_user_id) REFERENCES users(id)
        )
        """
    )

    # -- Lightweight SQLite migrations for older databases ---
    # ALL column additions MUST run BEFORE index creation
    _sqlite_add_column_if_missing(conn, "users", "portal_user_id", "INTEGER")
    _sqlite_add_column_if_missing(conn, "task_requests", "converted_at", "TEXT")
    for col, defn in (
        ("yapilan_isler", "TEXT"),
        ("gorev_ekleri", "TEXT"),
        ("harcama_bildirimleri", "TEXT"),
        ("gorev_tarih", "TEXT"),
        ("gorev_tarih_iso", "TEXT"),
        ("last_step", "INTEGER DEFAULT 0"),
        ("gorev_il", "TEXT"),
        ("gorev_ilce", "TEXT"),
        ("gorev_firma", "TEXT"),
        ("assigned_to_user_id", "INTEGER"),
        ("assigned_by_user_id", "INTEGER"),
        ("assigned_at", "TEXT"),
    ):
        _sqlite_add_column_if_missing(conn, "forms", col, defn)

    # -- Indexes (columns guaranteed to exist now) ---
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_forms_assigned_to ON forms(assigned_to_user_id)"
    )
    conn.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_users_portal_id "
        "ON users(portal_user_id) WHERE portal_user_id IS NOT NULL"
    )


def _sqlite_add_column_if_missing(
    conn: Connection, table: str, column: str, definition: str
) -> None:
    rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
    existing = {row["name"] for row in rows}
    if column not in existing:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def _ensure_schema_postgres(conn: Connection) -> None:
    """Create tables using PostgreSQL dialect."""

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            full_name TEXT NOT NULL,
            email TEXT,
            phone TEXT,
            password_hash TEXT,
            role TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            portal_user_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.execute(
        """
        CREATE UNIQUE INDEX IF NOT EXISTS idx_users_portal_id
            ON users(portal_user_id) WHERE portal_user_id IS NOT NULL
        """
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS forms (
            id SERIAL PRIMARY KEY,
            form_no TEXT NOT NULL UNIQUE,
            tarih TEXT,
            tarih_iso TEXT,
            dok_no TEXT,
            rev_no TEXT,
            avans TEXT,
            taseron TEXT,
            gorev_tanimi TEXT,
            gorev_yeri TEXT,
            gorev_yeri_lower TEXT,
            gorev_il TEXT,
            gorev_ilce TEXT,
            gorev_firma TEXT,
            gorev_tarih TEXT,
            gorev_tarih_iso TEXT,
            yapilan_isler TEXT,
            gorev_ekleri TEXT,
            harcama_bildirimleri TEXT,
            yola_cikis_tarih TEXT,
            yola_cikis_tarih_iso TEXT,
            yola_cikis_saat TEXT,
            donus_tarih TEXT,
            donus_tarih_iso TEXT,
            donus_saat TEXT,
            calisma_baslangic_tarih TEXT,
            calisma_baslangic_tarih_iso TEXT,
            calisma_baslangic_saat TEXT,
            calisma_bitis_tarih TEXT,
            calisma_bitis_tarih_iso TEXT,
            calisma_bitis_saat TEXT,
            mola_suresi TEXT,
            arac_plaka TEXT,
            hazirlayan TEXT,
            durum TEXT,
            personel_1 TEXT,
            personel_2 TEXT,
            personel_3 TEXT,
            personel_4 TEXT,
            personel_5 TEXT,
            personel_search TEXT,
            last_step INTEGER DEFAULT 0,
            assigned_to_user_id INTEGER,
            assigned_by_user_id INTEGER,
            assigned_at TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(assigned_to_user_id) REFERENCES users(id) ON DELETE SET NULL,
            FOREIGN KEY(assigned_by_user_id) REFERENCES users(id) ON DELETE SET NULL
        )
        """
    )
    conn.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_forms_form_no ON forms (form_no)"
    )
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_forms_assigned_to ON forms(assigned_to_user_id)"
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS form_sequence (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            last_no INTEGER NOT NULL DEFAULT 0
        )
        """
    )
    conn.execute(
        "INSERT INTO form_sequence (id, last_no) VALUES (1, 0) ON CONFLICT DO NOTHING"
    )

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS task_requests (
            id SERIAL PRIMARY KEY,
            customer_name TEXT NOT NULL,
            customer_phone TEXT,
            customer_email TEXT,
            customer_address TEXT,
            request_description TEXT NOT NULL,
            requirements TEXT,
            urgency TEXT DEFAULT 'normal',
            requested_by_user_id INTEGER NOT NULL REFERENCES users(id),
            status TEXT DEFAULT 'pending',
            notes TEXT,
            assigned_to_user_id INTEGER REFERENCES users(id),
            converted_form_no TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            converted_at TIMESTAMP
        )
        """
    )


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _sqlite_path(base_path: str) -> str:
    data_folder = os.environ.get("DATA_FOLDER", "").strip()
    if data_folder:
        os.makedirs(data_folder, exist_ok=True)
        return os.path.join(data_folder, DB_FILENAME)
    directory = base_path
    if directory:
        os.makedirs(directory, exist_ok=True)
    return os.path.join(base_path, DB_FILENAME)


def _convert_placeholders(query: str) -> str:
    """Convert SQLite ``?`` placeholders to PostgreSQL ``%s``."""
    return query.replace("?", "%s")
