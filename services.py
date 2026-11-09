"""Nghiệp vụ: dữ liệu bảng (kèm thuộc tính dẫn xuất), lương tài xế, doanh thu, bảo dưỡng xe."""
import math
from datetime import date, datetime, timedelta

import database
from utils import AppError

LUONG_CO_BAN_CHUYEN = 200_000      # VND / chuyến
HE_SO_LAI_XE = 2.0
HE_SO_PHU_XE = 1.0
CHU_KY_BAO_DUONG = 360             # ngày
KM_MOI_NGAY_GIAM = 100             # cứ 100 km quy đổi giảm 1 ngày


def _now_str(asof):
    return asof.strftime("%Y-%m-%d %H:%M")


# ---------------------------------------------------------------- bảo dưỡng
def maintenance(conn, asof=None):
    """Lịch bảo dưỡng của từng xe tại thời điểm asof (mặc định: hiện tại).

    Số km quy đổi = tổng (độ dài tuyến x hệ số đường khó) của các chuyến đã chạy kể từ ngày bảo dưỡng cuối.
    Số ngày còn lại = 360 - floor(km quy đổi / 100) - số ngày đã trôi qua.
    """
    asof = asof or datetime.now()
    today = asof.date()
    rows = conn.execute("""
        SELECT x.BienSo, x.HangSX || ' ' || x.Model AS TenXe, x.DoiXe, x.NgayBDCuoi,
               COALESCE(SUM(u.DoDaiKm * m.HeSoKmDuongKho), 0) AS KmQuyDoi
        FROM Xe x
        LEFT JOIN Chuyen c ON c.BienSo = x.BienSo
                          AND date(c.NgayKhoiHanh) >= x.NgayBDCuoi
                          AND c.NgayKhoiHanh <= ?
        LEFT JOIN Tuyen u ON u.MaTuyen = c.MaTuyen
        LEFT JOIN MucPhucTap m ON m.Muc = u.MucPhucTap
        GROUP BY x.BienSo ORDER BY x.BienSo""", (_now_str(asof),)).fetchall()
    result = []
    for r in rows:
        last = date.fromisoformat(r["NgayBDCuoi"])
        giam = int(math.floor(round(r["KmQuyDoi"], 6) / KM_MOI_NGAY_GIAM))
        han = CHU_KY_BAO_DUONG - giam
        try:
            ke_tiep = last + timedelta(days=han)
        except OverflowError:
            ke_tiep = date.min
        con_lai = han - (today - last).days
        result.append({"BienSo": r["BienSo"], "TenXe": r["TenXe"], "NgayBDCuoi": r["NgayBDCuoi"],
                       "SoNamSuDung": today.year - r["DoiXe"], "KmQuyDoi": r["KmQuyDoi"],
                       "NgayGiam": giam, "NgayBDKeTiep": ke_tiep.isoformat(),
                       "SoNgayConLai": con_lai, "QuaHan": con_lai <= 0})
    return result


def overdue_vehicles(conn, asof=None):
    """Danh sách xe quá hạn bảo dưỡng (số ngày còn lại <= 0)."""
    return [m for m in maintenance(conn, asof) if m["QuaHan"]]


# ---------------------------------------------------------------- lương tài xế
def salary_report(conn, year, month, asof=None):
    """Lương tháng (tích lũy đến thời điểm asof) của mọi tài xế.

    Lương mỗi chuyến = 200.000 x hệ số độ phức tạp của tuyến x hệ số vai trò (lái 2.0, phụ 1.0).
    """
    asof = asof or datetime.now()
    if not (1 <= month <= 12 and 1900 <= year <= 2200):
        raise AppError("Tháng/năm không hợp lệ.")
    trips = conn.execute("""
        SELECT c.MaLaiXe, c.MaPhuXe, m.HeSoLuong
        FROM Chuyen c
        JOIN Tuyen u ON u.MaTuyen = c.MaTuyen
        JOIN MucPhucTap m ON m.Muc = u.MucPhucTap
        WHERE substr(c.NgayKhoiHanh, 1, 7) = ? AND c.NgayKhoiHanh <= ?""",
        (f"{year:04d}-{month:02d}", _now_str(asof))).fetchall()
    acc = {}
    for t in trips:
        for ma, vai_tro, he_so in ((t["MaLaiXe"], 0, HE_SO_LAI_XE), (t["MaPhuXe"], 1, HE_SO_PHU_XE)):
            a = acc.setdefault(ma, [0, 0, 0.0])
            a[vai_tro] += 1
            a[2] += LUONG_CO_BAN_CHUYEN * t["HeSoLuong"] * he_so
    rows = []
    for d in conn.execute("SELECT MaTX, HoTen FROM TaiXe ORDER BY MaTX"):
        lai, phu, luong = acc.get(d["MaTX"], [0, 0, 0.0])
        rows.append([d["MaTX"], d["HoTen"], lai, phu, round(luong, 2)])
    return rows


# ---------------------------------------------------------------- doanh thu
def revenue_report(conn, d1, d2, asof=None):
    """Doanh thu từng xe trong khoảng [d1, d2] (chuỗi ISO): tổng (số khách x giá vé) các chuyến đã thực hiện."""
    asof = asof or datetime.now()
    if d1 > d2:
        raise AppError("Ngày bắt đầu phải nhỏ hơn hoặc bằng ngày kết thúc.")
    rows = conn.execute("""
        SELECT x.BienSo, x.HangSX || ' ' || x.Model, COUNT(c.MaChuyen),
               COALESCE(SUM(c.SoKhach), 0), COALESCE(SUM(c.SoKhach * c.GiaVe), 0)
        FROM Xe x
        LEFT JOIN Chuyen c ON c.BienSo = x.BienSo
                          AND date(c.NgayKhoiHanh) BETWEEN ? AND ?
                          AND c.NgayKhoiHanh <= ?
        GROUP BY x.BienSo ORDER BY x.BienSo""", (d1, d2, _now_str(asof))).fetchall()
    return [list(r) for r in rows]


# ---------------------------------------------------------------- dữ liệu bảng cho giao diện
def table_data(conn, ent):
    """Trả về (tiêu đề cột, kiểu cột, các dòng, thẻ tô màu) gồm cả các thuộc tính dẫn xuất."""
    headers = [f.label for f in ent.fields] + [d[1] for d in ent.derived]
    kinds = [f.kind for f in ent.fields] + [d[2] for d in ent.derived]
    today = date.today()
    info = {m["BienSo"]: m for m in maintenance(conn)} if ent.table == "Xe" else {}
    rows, tags = [], []
    for d in database.list_all(conn, ent):
        vals = [d[f.key] for f in ent.fields]
        tag = ""
        if ent.table == "Xe":
            m = info[d["BienSo"]]
            vals += [m["SoNamSuDung"], m["NgayBDKeTiep"], m["SoNgayConLai"]]
            tag = "overdue" if m["QuaHan"] else ""
        elif ent.table == "TaiXe":
            vals += [today.year - int(d["NgayVaoLam"][:4])]
        elif ent.table == "Chuyen":
            vals += [d["SoKhach"] * d["GiaVe"]]
        rows.append(vals)
        tags.append(tag)
    return headers, kinds, rows, tags
