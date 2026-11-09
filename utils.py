"""Hàm tiện ích: chuyển đổi kiểu dữ liệu và định dạng ngày/số cho người dùng Việt Nam."""
import math
from datetime import date, datetime

DATE_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y")
DATETIME_FORMATS = ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S",
                    "%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M")
TRUE_WORDS = {"1", "x", "co", "có", "yes", "y", "true", "t", "dung", "đúng"}
FALSE_WORDS = {"", "0", "no", "n", "false", "f", "khong", "không"}


class AppError(Exception):
    """Lỗi nhập liệu / nghiệp vụ, thông điệp đã được viết cho người dùng cuối."""


# ---------- chuyển đổi dữ liệu nhập (từ form hoặc Excel) ----------
def to_text(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def to_float(v):
    if isinstance(v, bool):
        raise ValueError("bool")
    if isinstance(v, (int, float)):
        f = float(v)
    else:
        s = str(v).strip().replace(" ", "")
        if "," in s:
            s = s.replace(",", "") if ("." in s or s.count(",") > 1) else s.replace(",", ".")
        f = float(s)
    if not math.isfinite(f):
        raise ValueError("not finite")
    return f


def to_int(v):
    f = to_float(v)
    if not f.is_integer():
        raise ValueError("not integer")
    return int(f)


def to_bool(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return bool(v)
    s = str(v).strip().lower()
    if s in TRUE_WORDS:
        return True
    if s in FALSE_WORDS:
        return False
    raise ValueError("bool")


def parse_date(v):
    """Nhận date/datetime của Excel hoặc chuỗi (yyyy-mm-dd, dd/mm/yyyy...) -> 'YYYY-MM-DD'."""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d")
        except ValueError:
            pass
    raise ValueError("date")


def parse_datetime(v):
    """Như parse_date nhưng có giờ phút -> 'YYYY-MM-DD HH:MM' (không có giờ thì lấy 00:00)."""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d") + " 00:00"
    s = str(v).strip()
    for f in DATETIME_FORMATS:
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d %H:%M")
        except ValueError:
            pass
    return parse_date(s) + " 00:00"


# ---------- định dạng hiển thị ----------
def fmt_date(iso):
    try:
        return datetime.strptime(str(iso)[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return str(iso)


def fmt_datetime(iso):
    try:
        return datetime.strptime(str(iso)[:16], "%Y-%m-%d %H:%M").strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return str(iso)


def fmt_float(v):
    s = f"{v:,.2f}"
    return s.rstrip("0").rstrip(".")


def fmt_value(kind, v):
    """Định dạng giá trị để hiển thị trong bảng."""
    if v is None:
        return ""
    if kind == "date":
        return fmt_date(v)
    if kind == "datetime":
        return fmt_datetime(v)
    if kind == "bool":
        return "Có" if v else ""
    if kind == "int":
        return str(int(v))
    if kind == "float":
        return fmt_float(float(v))
    return str(v)


def edit_text(kind, v):
    """Giá trị đưa vào ô nhập khi sửa (không có dấu phân cách hàng nghìn)."""
    if v is None:
        return ""
    if kind == "float":
        v = float(v)
        return str(int(v)) if v.is_integer() else str(v)
    return fmt_value(kind, v)


def sort_key(v):
    if v is None:
        return (0, 0)
    if isinstance(v, (int, float)):
        return (1, float(v))
    return (2, str(v).lower())
