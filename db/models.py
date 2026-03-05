from .database import get_connection


# ── Projects ─────────────────────────────────────────────────────────────────

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
                          hole_id: str, meters: float, standby: float) -> int:
    with get_connection() as conn:
        if entry_id:
            conn.execute(
                "UPDATE drilling_entries SET hole_id=?, meters_drilled=?, standby_hours=? WHERE id=?",
                (hole_id, meters, standby, entry_id)
            )
            conn.commit()
            return entry_id
        else:
            cur = conn.execute(
                "INSERT INTO drilling_entries (contractor_id, month, year, hole_id, meters_drilled, standby_hours) VALUES (?,?,?,?,?,?)",
                (contractor_id, month, year, hole_id, meters, standby)
            )
            conn.commit()
            return cur.lastrowid


def delete_drilling_entry(entry_id: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM drilling_entries WHERE id=?", (entry_id,))
        conn.commit()


# ── PPE Charges ───────────────────────────────────────────────────────────────

def get_ppe_charges(contractor_id: int, month: int, year: int):
    with get_connection() as conn:
        return conn.execute(
            "SELECT * FROM ppe_charges WHERE contractor_id=? AND month=? AND year=?",
            (contractor_id, month, year)
        ).fetchall()


def upsert_ppe_charge(charge_id: int | None, contractor_id: int, month: int, year: int,
                      item_name: str, quantity: float, unit_price: float) -> int:
    with get_connection() as conn:
        if charge_id:
            conn.execute(
                "UPDATE ppe_charges SET item_name=?, quantity=?, unit_price=? WHERE id=?",
                (item_name, quantity, unit_price, charge_id)
            )
            conn.commit()
            return charge_id
        else:
            cur = conn.execute(
                "INSERT INTO ppe_charges (contractor_id, month, year, item_name, quantity, unit_price) VALUES (?,?,?,?,?,?)",
                (contractor_id, month, year, item_name, quantity, unit_price)
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
