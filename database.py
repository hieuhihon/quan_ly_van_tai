"""Tầng truy cập dữ liệu: kết nối SQLite, kiểm tra dữ liệu nhập và CRUD tổng quát."""
import re
import sqlite3
from pathlib import Path

from models import ENTITIES
from utils import (AppError, to_text, to_int, to_float, to_bool, parse_date, parse_datetime)

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data" / "vantai.db"
SCHEMA_PATH = BASE_DIR / "schema.sql"

CONVERTERS = {"text": to_text, "int": to_int, "float": to_float, "bool": to_bool,
              "date": parse_date, "datetime": parse_datetime}
KIND_HINT = {"text": "văn bản", "int": "số nguyên", "float": "một số", "bool": "Có/Không",
             "date": "ngày dd/mm/yyyy", "datetime": "ngày giờ dd/mm/yyyy hh:mm"}

# Cột "ảo" (vai trò lái xe / phụ xe) được suy ra từ bảng chuyên biệt hóa
VIRTUAL_SQL = {
    "LaiXe": "EXISTS (SELECT 1 FROM LaiXe v WHERE v.MaTX = TaiXe.MaTX)",
    "PhuXe": "EXISTS (SELECT 1 FROM PhuXe v WHERE v.MaTX = TaiXe.MaTX)",
}

# Danh mục cho các ô chọn; giá trị hiển thị có dạng "<mã> - <mô tả>"
LOOKUPS = {
    "xe": "SELECT BienSo || ' - ' || HangSX || ' ' || Model || ' (' || SoGhe || ' ghế)' FROM Xe ORDER BY BienSo",
    "tuyen": "SELECT MaTuyen || ' - ' || DiemDau || ' → ' || DiemCuoi FROM Tuyen ORDER BY MaTuyen",
    "laixe": "SELECT t.MaTX || ' - ' || t.HoTen FROM TaiXe t JOIN LaiXe l ON l.MaTX = t.MaTX ORDER BY t.MaTX",
    "phuxe": "SELECT t.MaTX || ' - ' || t.HoTen FROM TaiXe t JOIN PhuXe p ON p.MaTX = t.MaTX ORDER BY t.MaTX",
    "muc": "SELECT Muc || ' - Cấp ' || Muc FROM MucPhucTap ORDER BY Muc",
}

CHECK_MESSAGES = {
    "ck_muc": "R3: Mức phức tạp chỉ được là 1, 2 hoặc 3.",
    "ck_muc_tuyen": "R3: Mức phức tạp chỉ được là 1, 2 hoặc 3.",
    "ck_heso": "Các hệ số phải lớn hơn 0.",
    "ck_soghe": "Số ghế phải lớn hơn 0.",
    "ck_doixe": "Đời xe không hợp lệ.",
    "ck_dodai": "Độ dài tuyến phải lớn hơn 0.",
    "ck_khach": "R5: Số khách phải lớn hơn hoặc bằng 0.",
    "ck_gia": "R5: Giá vé phải lớn hơn 0.",
    "ck_lai_phu": "R1: Lái xe và phụ xe trên cùng một chuyến phải khác nhau.",
}
LABELS = {f.key: f.label for e in ENTITIES for f in e.fields}


def connect(path=DB_PATH):
    """Mở (hoặc tạo mới) CSDL và đảm bảo cấu trúc bảng đã có."""
    if str(path) != ":memory:":
        Path(path).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    init_schema(conn)
    return conn


def init_schema(conn):
    conn.executescript(SCHEMA_PATH.read_text(encoding="utf-8"))


def friendly_error(e):
    """Đổi thông báo lỗi kỹ thuật của SQLite thành tiếng Việt dễ hiểu."""
    msg = str(e)
    if re.match(r"R\d:", msg):
        return msg
    m = re.search(r"CHECK constraint failed: (\w+)", msg)
    if m:
        return CHECK_MESSAGES.get(m.group(1), "Dữ liệu vi phạm ràng buộc: " + m.group(1))
    m = re.search(r"UNIQUE constraint failed: \w+\.(\w+)", msg)
    if m:
        return f"Giá trị của '{LABELS.get(m.group(1), m.group(1))}' đã tồn tại, không được trùng."
    m = re.search(r"NOT NULL constraint failed: \w+\.(\w+)", msg)
    if m:
        return f"Vui lòng nhập '{LABELS.get(m.group(1), m.group(1))}'."
    if "FOREIGN KEY" in msg:
        return ("Không thực hiện được: dữ liệu liên quan không tồn tại, hoặc bản ghi đang được sử dụng "
                "ở bảng khác (ví dụ: xe / tài xế / tuyến đã có chuyến).")
    return msg


def clean(ent, raw):
    """Kiểm tra bắt buộc và chuyển đổi kiểu cho dữ liệu nhập; trả về dict giá trị chuẩn hóa."""
    out = {}
    for f in ent.fields:
        if f.virtual and f.key not in raw:
            continue
        v = raw.get(f.key)
        if v is None or (isinstance(v, str) and not v.strip()):
            if f.required:
                raise AppError(f"Vui lòng nhập '{f.label}'.")
            out[f.key] = False if f.kind == "bool" else None
            continue
        try:
            out[f.key] = CONVERTERS[f.kind](v)
        except (ValueError, TypeError):
            raise AppError(f"Giá trị '{v}' của '{f.label}' không hợp lệ (cần nhập {KIND_HINT[f.kind]}).")
    return out


def _real(ent):
    return [f.key for f in ent.fields if not f.virtual]


def _select(ent):
    parts = _real(ent) + [f"{VIRTUAL_SQL[f.key]} AS {f.key}" for f in ent.fields if f.virtual]
    return ", ".join(parts)


def get(conn, ent, pk):
    row = conn.execute(f"SELECT {_select(ent)} FROM {ent.table} WHERE {ent.pk} = ?", (pk,)).fetchone()
    return dict(row) if row else None


def list_all(conn, ent):
    rows = conn.execute(f"SELECT {_select(ent)} FROM {ent.table} ORDER BY {ent.pk}").fetchall()
    return [dict(r) for r in rows]


def lookup(conn, name):
    return [r[0] for r in conn.execute(LOOKUPS[name])]


def _sync_roles(conn, ent, v):
    if ent.table != "TaiXe":
        return
    for table in ("LaiXe", "PhuXe"):
        if table not in v:
            continue
        if v[table]:
            conn.execute(f"INSERT OR IGNORE INTO {table} (MaTX) VALUES (?)", (v["MaTX"],))
        else:
            conn.execute(f"DELETE FROM {table} WHERE MaTX = ?", (v["MaTX"],))


def _insert(conn, ent, v):
    cols = _real(ent)
    sql = f"INSERT INTO {ent.table} ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})"
    try:
        with conn:
            conn.execute(sql, [v[c] for c in cols])
            _sync_roles(conn, ent, v)
    except sqlite3.Error as e:
        raise AppError(friendly_error(e)) from e


def _update(conn, ent, v):
    sets = [c for c in _real(ent) if c != ent.pk]
    sql = f"UPDATE {ent.table} SET {', '.join(c + ' = ?' for c in sets)} WHERE {ent.pk} = ?"
    try:
        with conn:
            cur = conn.execute(sql, [v[c] for c in sets] + [v[ent.pk]])
            if cur.rowcount == 0:
                raise AppError("Không tìm thấy bản ghi cần cập nhật.")
            _sync_roles(conn, ent, v)
    except sqlite3.Error as e:
        raise AppError(friendly_error(e)) from e


def insert(conn, ent, raw):
    _insert(conn, ent, clean(ent, raw))


def update(conn, ent, raw):
    _update(conn, ent, clean(ent, raw))


def upsert(conn, ent, raw):
    """Thêm mới nếu chưa có khóa chính, ngược lại cập nhật. Trả về 'them' hoặc 'sua'."""
    v = clean(ent, raw)
    if get(conn, ent, v[ent.pk]) is None:
        _insert(conn, ent, v)
        return "them"
    _update(conn, ent, v)
    return "sua"


def delete(conn, ent, pk):
    try:
        with conn:
            cur = conn.execute(f"DELETE FROM {ent.table} WHERE {ent.pk} = ?", (pk,))
    except sqlite3.Error as e:
        raise AppError(friendly_error(e)) from e
    if cur.rowcount == 0:
        raise AppError("Không tìm thấy bản ghi cần xóa.")


def clear_all(conn):
    """Xóa toàn bộ dữ liệu (giữ cấu trúc bảng, khôi phục hệ số mặc định của Mức phức tạp)."""
    with conn:
        for t in ("Chuyen", "LaiXe", "PhuXe", "TaiXe", "Xe", "Tuyen", "MucPhucTap"):
            conn.execute(f"DELETE FROM {t}")
    init_schema(conn)
