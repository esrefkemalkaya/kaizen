from .database import get_connection


# ── Projects ──────────────────────────────────────────────────────────────────

def get_projects():
    with get_connection() as conn:
        return conn.execute("SELECT * FROM projects ORDER BY name").fetchall()


def add_project(name: str, location: str):
    with get_connection() as conn:
        conn.execute("INSERT INTO projects (name, location) VALUES (?, ?)", (name, location))
        conn.commit()


def update_project(project_id: int, name: str, location: str):
    with get_connection() as conn:
        conn.execute("UPDATE projects SET name=?, location=? WHERE id=?", (name, location, project_id))
        conn.commit()


def delete_project(project_id: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM projects WHERE id=?", (project_id,))
        conn.commit()


# ── Contractors ───────────────────────────────────────────────────────────────

def get_contractors(project_id: int):
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM contractors WHERE project_id=? ORDER BY name",
            (project_id,)
        ).fetchall()


def get_contractor(contractor_id: int):
    with get_connection() as conn:
        return conn.execute("SELECT * FROM contractors WHERE id=?", (contractor_id,)).fetchone()


def add_contractor(project_id: int, name: str, ctype: str, rate_per_meter: float, standby_rate: float):
    with get_connection() as conn:
        conn.execute(
            "INSERT INTO contractors (project_id, name, type, rate_per_meter, standby_hour_rate) VALUES (?,?,?,?,?)",
            (project_id, name, ctype, rate_per_meter, standby_rate)
        )
        conn.commit()


def update_contractor(contractor_id: int, name: str, ctype: str, rate_per_meter: float, standby_rate: float):
    with get_connection() as conn:
        conn.execute(
            "UPDATE contractors SET name=?, type=?, rate_per_meter=?, standby_hour_rate=? WHERE id=?",
            (name, ctype, rate_per_meter, standby_rate, contractor_id)
        )
        conn.commit()


def delete_contractor(contractor_id: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM contractors WHERE id=?", (contractor_id,))
        conn.commit()


# ── Drilling Entries ──────────────────────────────────────────────────────────

def get_drilling_entries(contractor_id: int, month: int, year: int):
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM drilling_entries WHERE contractor_id=? AND month=? AND year=? ORDER BY hole_id",
            (contractor_id, month, year)
        ).fetchall()


def upsert_drilling_entry(entry_id: int | None, contractor_id: int, month: int, year: int,
                          hole_id: str, start_date: str, end_date: str,
                          start_depth: float, end_depth: float, meters: float) -> int:
    with get_connection() as conn:
        if entry_id:
            conn.execute(
                "UPDATE drilling_entries SET hole_id=?, start_date=?, end_date=?, "
                "start_depth=?, end_depth=?, meters_drilled=? WHERE id=?",
                (hole_id, start_date, end_date, start_depth, end_depth, meters, entry_id)
            )
            conn.commit()
            return entry_id
        else:
            cur = conn.execute(
                "INSERT INTO drilling_entries "
                "(contractor_id, month, year, hole_id, meters_drilled, standby_hours, "
                "start_date, end_date, start_depth, end_depth) "
                "VALUES (?,?,?,?,?,0,?,?,?,?)",
                (contractor_id, month, year, hole_id, meters,
                 start_date, end_date, start_depth, end_depth)
            )
            conn.commit()
            return cur.lastrowid


def delete_drilling_entry(entry_id: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM drilling_entries WHERE id=?", (entry_id,))
        conn.commit()


def get_all_hole_ids(contractor_id: int) -> list[str]:
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT DISTINCT hole_id FROM drilling_entries WHERE contractor_id=? ORDER BY hole_id",
            (contractor_id,)
        ).fetchall()
        return [r["hole_id"] for r in rows]


# ── Standby Entries ───────────────────────────────────────────────────────────

def get_standby_entries(contractor_id: int, month: int, year: int):
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM standby_entries "
            "WHERE contractor_id=? AND month=? AND year=? "
            "ORDER BY entry_date, start_time, id",
            (contractor_id, month, year)
        ).fetchall()


def upsert_standby_entry(entry_id: int | None, contractor_id: int, month: int, year: int,
                         entry_date: str, hole_id: str, start_time: str, end_time: str,
                         standby_type: str, description: str, hours: float,
                         rig_name: str = "") -> int:
    with get_connection() as conn:
        if entry_id:
            conn.execute(
                "UPDATE standby_entries SET entry_date=?, hole_id=?, start_time=?, end_time=?, "
                "standby_type=?, description=?, hours=?, rig_name=? WHERE id=?",
                (entry_date, hole_id, start_time, end_time,
                 standby_type, description, hours, rig_name, entry_id)
            )
            conn.commit()
            return entry_id
        else:
            cur = conn.execute(
                "INSERT INTO standby_entries "
                "(contractor_id, month, year, entry_date, hole_id, start_time, end_time, "
                "standby_type, description, hours, rig_name) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (contractor_id, month, year, entry_date, hole_id,
                 start_time, end_time, standby_type, description, hours, rig_name)
            )
            conn.commit()
            return cur.lastrowid


def delete_standby_entry(entry_id: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM standby_entries WHERE id=?", (entry_id,))
        conn.commit()


def calc_standby_rig_summary(entries, rate_s: float):
    """
    Apply blasting (Patlatma) deduction rule per rig:
      - Each rig gets 24 free Patlatma hours per month (routine underground operation).
      - Patlatma hours beyond 24 are payable like any other standby.
      - All non-Patlatma standby types are always payable in full.

    Returns:
        (rig_rows, total_payable_hours, total_amount)
        rig_rows: list of dicts per rig, sorted by rig name
    """
    from collections import defaultdict
    rigs = defaultdict(lambda: {"blasting": 0.0, "other": 0.0})
    for e in entries:
        rig   = (e["rig_name"] or "").strip() or "Unknown"
        hours = float(e["hours"] or 0)
        if (e["standby_type"] or "").strip().lower() == "patlatma":
            rigs[rig]["blasting"] += hours
        else:
            rigs[rig]["other"] += hours

    result = []
    total_payable_hours = 0.0
    total_amount = 0.0
    for rig in sorted(rigs):
        d = rigs[rig]
        blasting_total = d["blasting"]
        blasting_free  = min(blasting_total, 24.0)
        blasting_paid  = max(0.0, blasting_total - 24.0)
        other          = d["other"]
        payable_hours  = other + blasting_paid
        amount         = payable_hours * rate_s
        total_payable_hours += payable_hours
        total_amount        += amount
        result.append({
            "rig":            rig,
            "other":          other,
            "blasting_total": blasting_total,
            "blasting_free":  blasting_free,
            "blasting_paid":  blasting_paid,
            "payable_hours":  payable_hours,
            "amount":         amount,
        })
    return result, total_payable_hours, total_amount


def get_all_hole_ids_for_standby(contractor_id: int) -> list[str]:
    """All hole IDs seen in both drilling entries and standby entries."""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT DISTINCT hole_id FROM drilling_entries WHERE contractor_id=? "
            "UNION SELECT DISTINCT hole_id FROM standby_entries WHERE contractor_id=? AND hole_id != '' "
            "ORDER BY hole_id",
            (contractor_id, contractor_id)
        ).fetchall()
        return [r["hole_id"] for r in rows]


# ── PPE Charges ───────────────────────────────────────────────────────────────

def get_ppe_charges(contractor_id: int, month: int, year: int):
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM ppe_charges WHERE contractor_id=? AND month=? AND year=?",
            (contractor_id, month, year)
        ).fetchall()


def upsert_ppe_charge(charge_id: int | None, contractor_id: int, month: int, year: int,
                      item_name: str, quantity: float, unit_price: float,
                      material_code: str = "", unit_of_measure: str = "",
                      entry_date: str = "") -> int:
    with get_connection() as conn:
        if charge_id:
            conn.execute(
                """UPDATE ppe_charges
                   SET item_name=?, quantity=?, unit_price=?,
                       material_code=?, unit_of_measure=?, entry_date=?
                   WHERE id=?""",
                (item_name, quantity, unit_price,
                 material_code, unit_of_measure, entry_date, charge_id)
            )
            conn.commit()
            return charge_id
        else:
            cur = conn.execute(
                """INSERT INTO ppe_charges
                   (contractor_id, month, year, item_name, quantity, unit_price,
                    material_code, unit_of_measure, entry_date)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (contractor_id, month, year, item_name, quantity, unit_price,
                 material_code, unit_of_measure, entry_date)
            )
            conn.commit()
            return cur.lastrowid


def delete_ppe_charge(charge_id: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM ppe_charges WHERE id=?", (charge_id,))
        conn.commit()


# ── Diesel Charges ────────────────────────────────────────────────────────────

def get_diesel_charges(contractor_id: int, month: int, year: int):
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM diesel_charges WHERE contractor_id=? AND month=? AND year=?",
            (contractor_id, month, year)
        ).fetchall()


def upsert_diesel_charge(charge_id: int | None, contractor_id: int, month: int, year: int,
                         description: str, quantity: float, unit_price: float) -> int:
    with get_connection() as conn:
        if charge_id:
            conn.execute(
                "UPDATE diesel_charges SET description=?, quantity=?, unit_price=? WHERE id=?",
                (description, quantity, unit_price, charge_id)
            )
            conn.commit()
            return charge_id
        else:
            cur = conn.execute(
                "INSERT INTO diesel_charges (contractor_id, month, year, description, quantity, unit_price) VALUES (?,?,?,?,?,?)",
                (contractor_id, month, year, description, quantity, unit_price)
            )
            conn.commit()
            return cur.lastrowid


def delete_diesel_charge(charge_id: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM diesel_charges WHERE id=?", (charge_id,))
        conn.commit()


# ── Settings ──────────────────────────────────────────────────────────────────

def get_setting(key: str, default: str = '') -> str:
    with get_connection() as conn:
        row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default


def set_setting(key: str, value: str):
    with get_connection() as conn:
        conn.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES (?,?)", (key, value)
        )
        conn.commit()
