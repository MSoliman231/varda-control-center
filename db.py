# db.py (FULL) — Installable-safe version:
# - Uses %APPDATA%\VARDA Control Center\devices.db
# - Provides: init_db, register_device, list_stores, create_store, assign_device_to_store, export_devices_rows

import os
import sqlite3
from datetime import datetime
from pathlib import Path

MODEL_DEFAULT = "VR01"


def app_data_dir() -> str:
    base = os.getenv("APPDATA") or str(Path.home())
    p = Path(base) / "VARDA Control Center"
    p.mkdir(parents=True, exist_ok=True)
    return str(p)


def db_path() -> str:
    return os.path.join(app_data_dir(), "devices.db")


def connect():
    return sqlite3.connect(db_path())


def init_db():
    con = connect()
    try:
        cur = con.cursor()
        cur.execute("PRAGMA foreign_keys = ON;")

        cur.execute("""
        CREATE TABLE IF NOT EXISTS stores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            created_at TEXT DEFAULT (datetime('now'))
        );
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS devices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            serial TEXT NOT NULL UNIQUE,
            mac TEXT NOT NULL UNIQUE,
            model TEXT NOT NULL,
            batch TEXT,
            admin_password TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        """)

        cur.execute("""
        CREATE TABLE IF NOT EXISTS device_store (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            device_serial TEXT NOT NULL,
            store_id INTEGER NOT NULL,
            created_at TEXT DEFAULT (datetime('now')),
            UNIQUE(device_serial),
            FOREIGN KEY(store_id) REFERENCES stores(id) ON DELETE CASCADE
        );
        """)

        con.commit()
    finally:
        con.close()


def _next_serial(model: str) -> str:
    # Serial format: VR01-YYMM-0001
    now = datetime.now()
    yymm = now.strftime("%y%m")

    con = connect()
    try:
        cur = con.cursor()
        cur.execute("SELECT serial FROM devices WHERE serial LIKE ? ORDER BY serial DESC LIMIT 1;",
                    (f"{model}-{yymm}-%",))
        row = cur.fetchone()
        if not row:
            n = 1
        else:
            last = row[0]
            try:
                n = int(last.split("-")[-1]) + 1
            except Exception:
                n = 1
        return f"{model}-{yymm}-{n:04d}"
    finally:
        con.close()


def _generate_password(length=8) -> str:
    # simple numeric admin password
    import random
    return "".join(str(random.randint(0, 9)) for _ in range(length))


def register_device(model: str, mac: str):
    """
    Insert a new device with auto serial + generated admin password.
    Returns (serial, admin_password).
    """
    model = (model or MODEL_DEFAULT).strip()
    mac = (mac or "").strip().upper()
    if not mac:
        raise ValueError("MAC is required")

    serial = _next_serial(model)
    admin_password = _generate_password(8)

    con = connect()
    try:
        cur = con.cursor()
        cur.execute("""
            INSERT INTO devices (serial, mac, model, batch, admin_password)
            VALUES (?, ?, ?, ?, ?);
        """, (serial, mac, model, "", admin_password))
        con.commit()
        return serial, admin_password
    except sqlite3.IntegrityError as e:
        # MAC already exists or serial collision
        raise ValueError(f"Device already exists or MAC in use: {e}")
    finally:
        con.close()


def list_stores():
    con = connect()
    try:
        cur = con.cursor()
        cur.execute("SELECT id, name FROM stores ORDER BY name ASC;")
        return cur.fetchall()
    finally:
        con.close()


def create_store(name: str):
    name = (name or "").strip()
    if not name:
        raise ValueError("Store name is required")

    con = connect()
    try:
        cur = con.cursor()
        cur.execute("INSERT INTO stores (name) VALUES (?);", (name,))
        con.commit()
    except sqlite3.IntegrityError:
        raise ValueError("Store already exists")
    finally:
        con.close()


def assign_device_to_store(device_serial: str, store_id: int):
    device_serial = (device_serial or "").strip()
    if not device_serial:
        raise ValueError("Device serial is required")

    con = connect()
    try:
        cur = con.cursor()
        # upsert by UNIQUE(device_serial)
        cur.execute("""
            INSERT INTO device_store (device_serial, store_id)
            VALUES (?, ?)
            ON CONFLICT(device_serial) DO UPDATE SET store_id=excluded.store_id;
        """, (device_serial, int(store_id)))
        con.commit()
    finally:
        con.close()


def export_devices_rows():
    """
    Return rows with:
    serial, mac, model, batch, store_name, admin_password, created_at
    """
    con = connect()
    try:
        cur = con.cursor()
        cur.execute("""
            SELECT
                d.serial,
                d.mac,
                d.model,
                COALESCE(d.batch, ''),
                COALESCE(s.name, ''),
                COALESCE(d.admin_password, ''),
                d.created_at
            FROM devices d
            LEFT JOIN device_store ds ON ds.device_serial = d.serial
            LEFT JOIN stores s ON s.id = ds.store_id
            ORDER BY d.id DESC;
        """)
        return cur.fetchall()
    finally:
        con.close()