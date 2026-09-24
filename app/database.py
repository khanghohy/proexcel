import sqlite3
import os
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime

DB_DIR = Path(__file__).resolve().parent.parent / "data"
DB_PATH = DB_DIR / "excel_app.db"

def get_connection():
    DB_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    return conn

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        row_order INTEGER DEFAULT 0,
        name TEXT NOT NULL DEFAULT '',
        full_image_data BLOB,
        full_image_mime TEXT,
        full_image_name TEXT,
        full_image_size INTEGER DEFAULT 0,
        half_image_data BLOB,
        half_image_mime TEXT,
        half_image_name TEXT,
        half_image_size INTEGER DEFAULT 0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)
    conn.commit()

    # Check if empty, populate with 3 starter rows if brand new
    cursor.execute("SELECT COUNT(*) as count FROM records")
    count = cursor.fetchone()["count"]
    if count == 0:
        cursor.execute("INSERT INTO records (row_order, name) VALUES (1, 'Dự án Alpha')")
        cursor.execute("INSERT INTO records (row_order, name) VALUES (2, 'Kiểm tra giao diện')")
        cursor.execute("INSERT INTO records (row_order, name) VALUES (3, 'Báo cáo tuần')")
        conn.commit()

    conn.close()

def get_all_records() -> List[Dict[str, Any]]:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT 
            id, 
            row_order, 
            name, 
            full_image_name, 
            full_image_size, 
            full_image_mime,
            (full_image_data IS NOT NULL AND LENGTH(full_image_data) > 0) AS has_full_image,
            half_image_name, 
            half_image_size, 
            half_image_mime,
            (half_image_data IS NOT NULL AND LENGTH(half_image_data) > 0) AS has_half_image,
            created_at, 
            updated_at 
        FROM records 
        ORDER BY row_order ASC, id ASC
    """)
    rows = cursor.fetchall()
    result = []
    for r in rows:
        result.append({
            "id": r["id"],
            "row_order": r["row_order"],
            "name": r["name"],
            "full_image": {
                "exists": bool(r["has_full_image"]),
                "name": r["full_image_name"],
                "size": r["full_image_size"] or 0,
                "mime": r["full_image_mime"],
                "url": f"/api/records/{r['id']}/image/full" if r["has_full_image"] else None
            },
            "half_image": {
                "exists": bool(r["has_half_image"]),
                "name": r["half_image_name"],
                "size": r["half_image_size"] or 0,
                "mime": r["half_image_mime"],
                "url": f"/api/records/{r['id']}/image/half" if r["has_half_image"] else None
            },
            "created_at": r["created_at"],
            "updated_at": r["updated_at"]
        })
    conn.close()
    return result

def create_record(name: str = "") -> Dict[str, Any]:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(row_order) as max_order FROM records")
    row = cursor.fetchone()
    next_order = (row["max_order"] or 0) + 1
    
    cursor.execute("""
        INSERT INTO records (row_order, name)
        VALUES (?, ?)
    """, (next_order, name))
    conn.commit()
    new_id = cursor.lastrowid
    conn.close()
    return {
        "id": new_id,
        "row_order": next_order,
        "name": name,
        "full_image": {"exists": False, "name": None, "size": 0, "mime": None, "url": None},
        "half_image": {"exists": False, "name": None, "size": 0, "mime": None, "url": None},
        "created_at": datetime.now().isoformat(),
        "updated_at": datetime.now().isoformat()
    }

def update_record(record_id: int, name: str) -> bool:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE records 
        SET name = ?, updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
    """, (name, record_id))
    conn.commit()
    affected = cursor.rowcount > 0
    conn.close()
    return affected

def delete_record(record_id: int) -> bool:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM records WHERE id = ?", (record_id,))
    conn.commit()
    affected = cursor.rowcount > 0
    conn.close()
    return affected

def save_image_blob(record_id: int, image_type: str, data_bytes: bytes, filename: str, mime_type: str) -> bool:
    """Store image data directly as BLOB in the SQL database."""
    conn = get_connection()
    cursor = conn.cursor()
    size = len(data_bytes)
    
    if image_type == "full":
        cursor.execute("""
            UPDATE records 
            SET full_image_data = ?, 
                full_image_mime = ?, 
                full_image_name = ?, 
                full_image_size = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (data_bytes, mime_type, filename, size, record_id))
    elif image_type == "half":
        cursor.execute("""
            UPDATE records 
            SET half_image_data = ?, 
                half_image_mime = ?, 
                half_image_name = ?, 
                half_image_size = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (data_bytes, mime_type, filename, size, record_id))
    else:
        conn.close()
        return False
        
    conn.commit()
    affected = cursor.rowcount > 0
    conn.close()
    return affected

def remove_image_blob(record_id: int, image_type: str) -> bool:
    conn = get_connection()
    cursor = conn.cursor()
    if image_type == "full":
        cursor.execute("""
            UPDATE records 
            SET full_image_data = NULL, 
                full_image_mime = NULL, 
                full_image_name = NULL, 
                full_image_size = 0,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (record_id,))
    elif image_type == "half":
        cursor.execute("""
            UPDATE records 
            SET half_image_data = NULL, 
                half_image_mime = NULL, 
                half_image_name = NULL, 
                half_image_size = 0,
                updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (record_id,))
    else:
        conn.close()
        return False
        
    conn.commit()
    affected = cursor.rowcount > 0
    conn.close()
    return affected

def get_image_blob(record_id: int, image_type: str) -> Optional[Tuple[bytes, str, str]]:
    conn = get_connection()
    cursor = conn.cursor()
    if image_type == "full":
        cursor.execute("SELECT full_image_data, full_image_mime, full_image_name FROM records WHERE id = ?", (record_id,))
    elif image_type == "half":
        cursor.execute("SELECT half_image_data, half_image_mime, half_image_name FROM records WHERE id = ?", (record_id,))
    else:
        conn.close()
        return None
        
    row = cursor.fetchone()
    conn.close()
    if not row or not row[0]:
        return None
    return row[0], row[1] or "image/png", row[2] or f"{image_type}_{record_id}.png"

def get_database_info() -> Dict[str, Any]:
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) as total_rows FROM records")
    total_rows = cursor.fetchone()["total_rows"]
    
    cursor.execute("SELECT COUNT(*) as count FROM records WHERE full_image_data IS NOT NULL")
    full_count = cursor.fetchone()["count"]

    cursor.execute("SELECT COUNT(*) as count FROM records WHERE half_image_data IS NOT NULL")
    half_count = cursor.fetchone()["count"]

    # Table schema
    cursor.execute("PRAGMA table_info(records)")
    columns = [dict(c) for c in cursor.fetchall()]

    conn.close()

    db_size = os.path.getsize(DB_PATH) if DB_PATH.exists() else 0
    
    return {
        "engine": "SQLite 3 (Embedded SQL Server/Engine)",
        "db_path": str(DB_PATH),
        "db_size_bytes": db_size,
        "db_size_formatted": f"{db_size / (1024 * 1024):.2f} MB" if db_size > 1024*1024 else f"{db_size / 1024:.2f} KB",
        "total_rows": total_rows,
        "full_images_count": full_count,
        "half_images_count": half_count,
        "total_images_stored": full_count + half_count,
        "columns": columns
    }

def run_custom_sql(query: str) -> Dict[str, Any]:
    cleaned = query.strip()
    # For safety in query inspector, limit or handle SELECT
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute(cleaned)
        if cleaned.upper().startswith("SELECT") or cleaned.upper().startswith("PRAGMA"):
            rows = cursor.fetchall()
            col_names = [d[0] for d in cursor.description] if cursor.description else []
            # Serialize rows
            serialized = []
            for r in rows:
                row_dict = {}
                for k in col_names:
                    val = r[k]
                    if isinstance(val, bytes):
                        row_dict[k] = f"<BLOB {len(val)} bytes>"
                    else:
                        row_dict[k] = val
                serialized.append(row_dict)
            conn.close()
            return {"success": True, "columns": col_names, "rows": serialized, "rowcount": len(serialized)}
        else:
            conn.commit()
            affected = cursor.rowcount
            conn.close()
            return {"success": True, "affected": affected, "message": f"Query executed successfully, {affected} rows affected."}
    except Exception as e:
        conn.close()
        return {"success": False, "error": str(e)}
