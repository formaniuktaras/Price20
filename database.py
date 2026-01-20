"""Database access layer for catalog entities and specs."""
from __future__ import annotations

import logging
import sqlite3
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Sequence, Tuple, Literal, overload

from app_paths import get_db_path

LOGGER = logging.getLogger(__name__)

# Retain the symbol for backward compatibility, but prefer get_db_path().
DB_FILE = str(get_db_path())


def db_connect() -> sqlite3.Connection:
    db_path = get_db_path()
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _deduplicate_model_specs(cur: sqlite3.Cursor, conn: sqlite3.Connection) -> None:
    cur.execute(
        """
        SELECT model_id, key, MAX(id) AS keep_id, COUNT(*) AS cnt
        FROM model_specs
        GROUP BY model_id, key
        HAVING cnt > 1
        """
    )
    rows = cur.fetchall()
    total_removed = 0
    for model_id, key, keep_id, _cnt in rows:
        cur.execute(
            "DELETE FROM model_specs WHERE model_id=? AND key=? AND id<>?",
            (model_id, key, keep_id),
        )
        total_removed += cur.rowcount
    if total_removed:
        LOGGER.info("Removed %s duplicate model_specs rows before adding unique index", total_removed)
        conn.commit()


def init_db() -> None:
    conn = db_connect()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS categories(
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            name       TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS brands(
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            category_id INTEGER NOT NULL,
            name        TEXT NOT NULL,
            created_at  TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(category_id, name),
            FOREIGN KEY(category_id) REFERENCES categories(id) ON DELETE CASCADE
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS models(
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            brand_id INTEGER NOT NULL,
            name     TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(brand_id, name),
            FOREIGN KEY(brand_id) REFERENCES brands(id) ON DELETE CASCADE
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS model_specs(
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            model_id INTEGER NOT NULL,
            key      TEXT NOT NULL,
            value    TEXT,
            FOREIGN KEY(model_id) REFERENCES models(id) ON DELETE CASCADE
        )
        """
    )
    conn.commit()

    def ensure_created_at(table: str):
        cur.execute(f"PRAGMA table_info({table})")
        columns = {row[1] for row in cur.fetchall()}
        if "created_at" not in columns:
            cur.execute(
                f"ALTER TABLE {table} ADD COLUMN created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP"
            )
            conn.commit()

    ensure_created_at("categories")
    ensure_created_at("brands")
    ensure_created_at("models")

    _deduplicate_model_specs(cur, conn)
    cur.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_model_specs_model_key ON model_specs(model_id, key)"
    )
    conn.commit()

    cur.execute("SELECT COUNT(*) FROM categories")
    if cur.fetchone()[0] == 0:
        cur.executemany(
            "INSERT INTO categories(name) VALUES(?)",
            [("Смартфони",), ("Планшети",)],
        )
        conn.commit()
    conn.close()


# ---- CRUD helpers -----------------------------------------------------------------

def _trimmed_rows(rows: Sequence[Sequence[object]]) -> List[Tuple[object, ...]]:
    trimmed: List[Tuple[object, ...]] = []
    for row in rows:
        if not row:
            continue
        idx = row[0]
        name = row[1] if len(row) > 1 else None
        if isinstance(name, str):
            name = name.strip()
        if len(row) == 2:
            trimmed.append((idx, name))
        elif len(row) >= 3:
            trimmed.append((idx, name, *row[2:]))
        else:
            trimmed.append(tuple(row))
    return trimmed


@overload
def get_categories(include_created: Literal[False] = False) -> List[Tuple[int, str]]:
    ...


@overload
def get_categories(include_created: Literal[True]) -> List[Tuple[int, str, Optional[str]]]:
    ...


def get_categories(include_created: bool = False):
    conn = db_connect()
    cur = conn.cursor()
    if include_created:
        cur.execute("SELECT id, name, created_at FROM categories ORDER BY name")
    else:
        cur.execute("SELECT id, name FROM categories ORDER BY name")
    rows = cur.fetchall()
    conn.close()
    return _trimmed_rows(rows)


def add_category(name: str) -> None:
    name = name.strip()
    if not name:
        return
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("INSERT OR IGNORE INTO categories(name) VALUES(?)", (name,))
    conn.commit()
    conn.close()


def rename_category(cat_id: int, new_name: str):
    new_name = new_name.strip()
    if not new_name:
        return False
    conn = db_connect()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE categories SET name=? WHERE id=?", (new_name, cat_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        return exc
    finally:
        conn.close()


def delete_category(cat_id: int) -> None:
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM categories WHERE id=?", (cat_id,))
    conn.commit()
    conn.close()


@overload
def get_brands(category_id: int, include_created: Literal[False] = False) -> List[Tuple[int, str]]:
    ...


@overload
def get_brands(category_id: int, include_created: Literal[True]) -> List[Tuple[int, str, Optional[str]]]:
    ...


def get_brands(category_id: int, include_created: bool = False):
    conn = db_connect()
    cur = conn.cursor()
    if include_created:
        cur.execute(
            "SELECT id, name, created_at FROM brands WHERE category_id=? ORDER BY name",
            (category_id,),
        )
    else:
        cur.execute(
            "SELECT id, name FROM brands WHERE category_id=? ORDER BY name",
            (category_id,),
        )
    rows = cur.fetchall()
    conn.close()
    return _trimmed_rows(rows)


def add_brand(category_id: int, name: str) -> None:
    name = name.strip()
    if not name:
        return
    conn = db_connect()
    cur = conn.cursor()
    cur.execute(
        "INSERT OR IGNORE INTO brands(category_id, name) VALUES(?,?)",
        (category_id, name),
    )
    conn.commit()
    conn.close()


def rename_brand(brand_id: int, new_name: str):
    new_name = new_name.strip()
    if not new_name:
        return False
    conn = db_connect()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE brands SET name=? WHERE id=?", (new_name, brand_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        return exc
    finally:
        conn.close()


def delete_brand(brand_id: int) -> None:
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM brands WHERE id=?", (brand_id,))
    conn.commit()
    conn.close()


@overload
def get_models(brand_id: int, include_created: Literal[False] = False) -> List[Tuple[int, str]]:
    ...


@overload
def get_models(brand_id: int, include_created: Literal[True]) -> List[Tuple[int, str, Optional[str]]]:
    ...


def get_models(brand_id: int, include_created: bool = False):
    conn = db_connect()
    cur = conn.cursor()
    if include_created:
        cur.execute(
            "SELECT id, name, created_at FROM models WHERE brand_id=? ORDER BY name",
            (brand_id,),
        )
    else:
        cur.execute(
            "SELECT id, name FROM models WHERE brand_id=? ORDER BY name",
            (brand_id,),
        )
    rows = cur.fetchall()
    conn.close()
    return _trimmed_rows(rows)


def add_model(brand_id: int, name: str) -> None:
    name = name.strip()
    if not name:
        return
    conn = db_connect()
    cur = conn.cursor()
    cur.execute(
        "INSERT OR IGNORE INTO models(brand_id, name) VALUES(?,?)",
        (brand_id, name),
    )
    conn.commit()
    conn.close()


def rename_model(model_id: int, new_name: str):
    new_name = new_name.strip()
    if not new_name:
        return False
    conn = db_connect()
    cur = conn.cursor()
    try:
        cur.execute("UPDATE models SET name=? WHERE id=?", (new_name, model_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        return exc
    finally:
        conn.close()


def delete_model(model_id: int) -> None:
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM models WHERE id=?", (model_id,))
    conn.commit()
    conn.close()


# ---- Specs (key-value) -------------------------------------------------------------

def get_specs(model_id: int) -> List[Tuple[int, str, Optional[str]]]:
    conn = db_connect()
    cur = conn.cursor()
    cur.execute(
        "SELECT id, key, value FROM model_specs WHERE model_id=? ORDER BY id",
        (model_id,),
    )
    rows = cur.fetchall()
    conn.close()
    return rows


def insert_spec(model_id: int, key: str, value: str) -> Optional[int]:
    key = key.strip()
    if not key:
        return None
    conn = db_connect()
    cur = conn.cursor()
    try:
        cur.execute(
            """
            INSERT INTO model_specs(model_id, key, value) VALUES(?,?,?)
            ON CONFLICT(model_id, key) DO UPDATE SET value=excluded.value
            RETURNING id
            """,
            (model_id, key, value),
        )
        row = cur.fetchone()
        conn.commit()
        inserted_id = row[0] if row else None
        return inserted_id
    except sqlite3.OperationalError:
        cur.execute(
            """
            INSERT INTO model_specs(model_id, key, value) VALUES(?,?,?)
            ON CONFLICT(model_id, key) DO UPDATE SET value=excluded.value
            """,
            (model_id, key, value),
        )
        conn.commit()
        cur.execute(
            "SELECT id FROM model_specs WHERE model_id=? AND key=?",
            (model_id, key),
        )
        row = cur.fetchone()
        inserted_id = row[0] if row else None
        conn.commit()
        return inserted_id
    finally:
        conn.close()


def update_spec(spec_id: int, key: str, value: str) -> None:
    key = key.strip()
    if not key:
        return
    conn = db_connect()
    cur = conn.cursor()
    try:
        cur.execute("SELECT model_id FROM model_specs WHERE id=?", (spec_id,))
        row = cur.fetchone()
        if row is None:
            return
        model_id = row[0]
        cur.execute(
            "SELECT 1 FROM model_specs WHERE model_id=? AND key=? AND id<>?",
            (model_id, key, spec_id),
        )
        if cur.fetchone():
            raise ValueError("Параметр з таким ключем вже існує для цієї моделі.")
        cur.execute(
            "UPDATE model_specs SET key=?, value=? WHERE id=?",
            (key, value, spec_id),
        )
        conn.commit()
    except sqlite3.IntegrityError as exc:
        conn.rollback()
        raise ValueError("Не вдалося оновити характеристику: ключ має бути унікальним для моделі.") from exc
    finally:
        conn.close()


def delete_spec(spec_id: int) -> None:
    conn = db_connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM model_specs WHERE id=?", (spec_id,))
    conn.commit()
    conn.close()


def replace_specs(model_id: int, specs: Sequence[Tuple[str, str]]) -> None:
    """Replace all specifications for a model while preserving order."""

    conn = db_connect()
    cur = conn.cursor()
    cur.execute("DELETE FROM model_specs WHERE model_id=?", (model_id,))
    if specs:
        payload: List[Tuple[int, str, str]] = []
        for key, value in specs:
            normalized_key = key.strip()
            if not normalized_key:
                continue
            payload.append((model_id, normalized_key, value))
        if payload:
            cur.executemany(
                "INSERT INTO model_specs(model_id, key, value) VALUES(?,?,?)",
                payload,
            )
    conn.commit()
    conn.close()


def load_specs_map(model_ids: Iterable[int]) -> Dict[int, Dict[str, Optional[str]]]:
    unique_ids: List[int] = []
    seen = set()
    for mid in model_ids:
        try:
            ivalue = int(mid)
        except (TypeError, ValueError):
            continue
        if ivalue in seen:
            continue
        seen.add(ivalue)
        unique_ids.append(ivalue)
    if not unique_ids:
        return {}

    def _chunks(seq, size):
        for i in range(0, len(seq), size):
            yield seq[i : i + size]

    specs_map: Dict[int, Dict[str, Optional[str]]] = {}

    conn = db_connect()
    cur = conn.cursor()
    for chunk in _chunks(unique_ids, 500):
        placeholders = ",".join(["?"] * len(chunk))
        query = f"""
            SELECT model_id, key, value
            FROM model_specs
            WHERE model_id IN ({placeholders})
            ORDER BY model_id, id
        """
        cur.execute(query, tuple(chunk))
        rows = cur.fetchall()

        for model_id, key, value in rows:
            if isinstance(key, str):
                key = key.strip()
            if isinstance(value, str):
                value = value.strip()
            if not key:
                continue
            specs_map.setdefault(model_id, {})[key] = value
    conn.close()
    return specs_map


def export_catalog_dump() -> Dict[str, List[Dict[str, object]]]:
    """Return the complete catalog dataset for backup/export purposes."""

    conn = db_connect()
    cur = conn.cursor()

    cur.execute("SELECT id, name, created_at FROM categories ORDER BY id")
    categories = [
        {
            "id": row[0],
            "name": (row[1] or "").strip() if isinstance(row[1], str) else row[1],
            "created_at": row[2],
        }
        for row in cur.fetchall()
        if row and row[0]
    ]

    cur.execute("SELECT id, category_id, name, created_at FROM brands ORDER BY id")
    brands = [
        {
            "id": row[0],
            "category_id": row[1],
            "name": (row[2] or "").strip() if isinstance(row[2], str) else row[2],
            "created_at": row[3],
        }
        for row in cur.fetchall()
        if row and row[0]
    ]

    cur.execute("SELECT id, brand_id, name, created_at FROM models ORDER BY id")
    models = [
        {
            "id": row[0],
            "brand_id": row[1],
            "name": (row[2] or "").strip() if isinstance(row[2], str) else row[2],
            "created_at": row[3],
        }
        for row in cur.fetchall()
        if row and row[0]
    ]

    cur.execute("SELECT id, model_id, key, value FROM model_specs ORDER BY id")
    specs = [
        {
            "id": row[0],
            "model_id": row[1],
            "key": (row[2] or "").strip() if isinstance(row[2], str) else row[2],
            "value": (row[3] or "").strip() if isinstance(row[3], str) else row[3],
        }
        for row in cur.fetchall()
        if row and row[0]
    ]

    conn.close()

    return {
        "categories": categories,
        "brands": brands,
        "models": models,
        "specs": specs,
    }


def import_catalog_dump(payload: Dict[str, Iterable[Dict[str, object]]]) -> None:
    """Replace the catalog dataset with values from an exported dump."""

    if not isinstance(payload, dict):
        raise ValueError("Invalid catalog payload: expected a dictionary")

    categories_raw = payload.get("categories") or []
    brands_raw = payload.get("brands") or []
    models_raw = payload.get("models") or []
    specs_raw = payload.get("specs") or []

    def _normalize_created(value):
        if isinstance(value, str) and value.strip():
            return value.strip()
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _to_int(value):
        try:
            number = int(value)
        except (TypeError, ValueError):
            return None
        return number if number >= 0 else None

    categories: List[Tuple[int, str, str]] = []
    seen_category_ids = set()
    for item in categories_raw:
        if isinstance(item, dict):
            raw_id = item.get("id")
            name = item.get("name")
            created_at = item.get("created_at")
        elif isinstance(item, (list, tuple)) and len(item) >= 3:
            raw_id, name, created_at = item[0], item[1], item[2]
        else:
            raise ValueError("Invalid category entry in catalog payload")
        cat_id = _to_int(raw_id)
        if not cat_id:
            raise ValueError("Category ID must be a positive integer")
        if cat_id in seen_category_ids:
            raise ValueError(f"Duplicate category ID: {cat_id}")
        label = (name or "").strip() if isinstance(name, str) else name
        if not label:
            raise ValueError(f"Category {cat_id} is missing a name")
        categories.append((cat_id, label, _normalize_created(created_at)))
        seen_category_ids.add(cat_id)

    brands: List[Tuple[int, int, str, str]] = []
    seen_brand_ids = set()
    valid_categories = seen_category_ids
    for item in brands_raw:
        if isinstance(item, dict):
            raw_id = item.get("id")
            raw_cat = item.get("category_id")
            name = item.get("name")
            created_at = item.get("created_at")
        elif isinstance(item, (list, tuple)) and len(item) >= 4:
            raw_id, raw_cat, name, created_at = item[0], item[1], item[2], item[3]
        else:
            raise ValueError("Invalid brand entry in catalog payload")
        brand_id = _to_int(raw_id)
        if not brand_id:
            raise ValueError("Brand ID must be a positive integer")
        if brand_id in seen_brand_ids:
            raise ValueError(f"Duplicate brand ID: {brand_id}")
        category_id = _to_int(raw_cat)
        if not category_id or category_id not in valid_categories:
            raise ValueError(f"Brand {brand_id} references unknown category {raw_cat}")
        label = (name or "").strip() if isinstance(name, str) else name
        if not label:
            raise ValueError(f"Brand {brand_id} is missing a name")
        brands.append((brand_id, category_id, label, _normalize_created(created_at)))
        seen_brand_ids.add(brand_id)

    models: List[Tuple[int, int, str, str]] = []
    seen_model_ids = set()
    valid_brands = seen_brand_ids
    for item in models_raw:
        if isinstance(item, dict):
            raw_id = item.get("id")
            raw_brand = item.get("brand_id")
            name = item.get("name")
            created_at = item.get("created_at")
        elif isinstance(item, (list, tuple)) and len(item) >= 4:
            raw_id, raw_brand, name, created_at = item[0], item[1], item[2], item[3]
        else:
            raise ValueError("Invalid model entry in catalog payload")
        model_id = _to_int(raw_id)
        if not model_id:
            raise ValueError("Model ID must be a positive integer")
        if model_id in seen_model_ids:
            raise ValueError(f"Duplicate model ID: {model_id}")
        brand_id = _to_int(raw_brand)
        if not brand_id or brand_id not in valid_brands:
            raise ValueError(f"Model {model_id} references unknown brand {raw_brand}")
        label = (name or "").strip() if isinstance(name, str) else name
        if not label:
            raise ValueError(f"Model {model_id} is missing a name")
        models.append((model_id, brand_id, label, _normalize_created(created_at)))
        seen_model_ids.add(model_id)

    specs: List[Tuple[int, int, str, Optional[str]]] = []
    valid_models = seen_model_ids
    for item in specs_raw:
        if isinstance(item, dict):
            raw_id = item.get("id")
            raw_model = item.get("model_id")
            key = item.get("key")
            value = item.get("value")
        elif isinstance(item, (list, tuple)) and len(item) >= 4:
            raw_id, raw_model, key, value = item[0], item[1], item[2], item[3]
        else:
            raise ValueError("Invalid spec entry in catalog payload")
        spec_id = _to_int(raw_id)
        if not spec_id:
            raise ValueError("Specification ID must be a positive integer")
        model_id = _to_int(raw_model)
        if not model_id or model_id not in valid_models:
            raise ValueError(f"Specification {spec_id} references unknown model {raw_model}")
        normalized_key = (key or "").strip() if isinstance(key, str) else key
        if not normalized_key:
            raise ValueError(f"Specification {spec_id} is missing a key")
        normalized_value = (value or "").strip() if isinstance(value, str) else value
        specs.append((spec_id, model_id, normalized_key, normalized_value))

    conn = db_connect()
    cur = conn.cursor()
    try:
        cur.execute("BEGIN")
        cur.execute("DELETE FROM model_specs")
        cur.execute("DELETE FROM models")
        cur.execute("DELETE FROM brands")
        cur.execute("DELETE FROM categories")
        cur.execute(
            "DELETE FROM sqlite_sequence WHERE name IN ('model_specs','models','brands','categories')"
        )

        if categories:
            cur.executemany(
                "INSERT INTO categories(id, name, created_at) VALUES(?,?,?)",
                categories,
            )
        if brands:
            cur.executemany(
                "INSERT INTO brands(id, category_id, name, created_at) VALUES(?,?,?,?)",
                brands,
            )
        if models:
            cur.executemany(
                "INSERT INTO models(id, brand_id, name, created_at) VALUES(?,?,?,?)",
                models,
            )
        if specs:
            cur.executemany(
                "INSERT INTO model_specs(id, model_id, key, value) VALUES(?,?,?,?)",
                specs,
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _normalize_id_list(ids) -> List[int]:
    if not ids:
        return []
    if isinstance(ids, (list, tuple, set)):
        result: List[int] = []
        for value in ids:
            try:
                ivalue = int(value)
            except (TypeError, ValueError):
                continue
            if ivalue:
                result.append(ivalue)
        return result
    try:
        ivalue = int(ids)
    except (TypeError, ValueError):
        return []
    return [ivalue] if ivalue else []


def collect_models(
    category_ids=None,
    brand_ids=None,
    model_ids=None,
) -> List[Tuple[str, str, str, int, int, int]]:
    category_ids = _normalize_id_list(category_ids)
    brand_ids = _normalize_id_list(brand_ids)
    model_ids = _normalize_id_list(model_ids)

    conn = db_connect()
    cur = conn.cursor()
    query = """
        SELECT b.name, m.name, c.name, m.id, b.id, c.id
        FROM models m
        JOIN brands b ON m.brand_id = b.id
        JOIN categories c ON b.category_id = c.id
    """
    conditions: List[str] = []
    params: List[int] = []
    if category_ids:
        placeholders = ",".join(["?"] * len(category_ids))
        conditions.append(f"c.id IN ({placeholders})")
        params.extend(category_ids)
    if brand_ids:
        placeholders = ",".join(["?"] * len(brand_ids))
        conditions.append(f"b.id IN ({placeholders})")
        params.extend(brand_ids)
    if model_ids:
        placeholders = ",".join(["?"] * len(model_ids))
        conditions.append(f"m.id IN ({placeholders})")
        params.extend(model_ids)
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    query += " ORDER BY c.name, b.name, m.name"
    cur.execute(query, tuple(params))
    rows = cur.fetchall()
    conn.close()
    cleaned: List[Tuple[str, str, str, int, int, int]] = []
    for brand, model, cat, mid, brand_id, cat_id in rows:
        if isinstance(brand, str):
            brand = brand.strip()
        if isinstance(model, str):
            model = model.strip()
        if isinstance(cat, str):
            cat = cat.strip()
        cleaned.append((brand, model, cat, mid, brand_id, cat_id))
    return cleaned
