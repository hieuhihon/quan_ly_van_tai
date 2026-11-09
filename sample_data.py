"""Dữ liệu mẫu (tạo theo ngày hiện tại để báo cáo luôn có số liệu) dùng cho file Excel mẫu và chế độ demo."""
from datetime import date, datetime, time, timedelta


def build(today=None):
    """Trả về dict {tên sheet: danh sách dòng}, thứ tự cột đúng với models.py."""
    today = today or date.today()

    def ago(n):
        return today - timedelta(days=n)

    def at(n, hour):
        return datetime.combine(ago(n), time(hour, 0))

    muc = [(1, 1.0, 1.0), (2, 1.2, 1.5), (3, 1.5, 2.0)]
    tuyen = [
        ("T01", "Hà Nội", "Hải Phòng", 120, 1),
        ("T02", "Hà Nội", "Đà Nẵng", 770, 2),
        ("T03", "Hà Nội", "TP. Hồ Chí Minh", 1700, 3),
        ("T04", "Hà Nội", "Sa Pa", 320, 3),
        ("T05", "Đà Nẵng", "TP. Hồ Chí Minh", 960, 2),
    ]
    xe = [
        ("29B-123.45", "Trắng", "Hyundai", 2021, "Universe Express", 45, ago(120)),
        ("29B-234.56", "Xanh", "Thaco", 2019, "Mobihome", 34, ago(330)),
        ("30F-345.67", "Đỏ", "Samco", 2020, "Felix", 29, ago(200)),
        ("29B-456.78", "Trắng", "Hyundai", 2022, "Solati", 16, ago(60)),
        ("30H-567.89", "Xanh", "Thaco", 2018, "Garden", 45, ago(400)),
    ]
    taixe = [
        ("TX01", "Nguyễn Văn An", "001085012345", "790012345", "E", "Cầu Giấy, Hà Nội", date(1985, 3, 12), date(2016, 5, 1), 1, 0),
        ("TX02", "Trần Văn Bình", "001082023456", "790023456", "FC", "Thanh Xuân, Hà Nội", date(1982, 7, 25), date(2015, 9, 15), 1, 0),
        ("TX03", "Lê Quang Cường", "001088034567", "790034567", "E", "Hà Đông, Hà Nội", date(1988, 11, 2), date(2018, 1, 10), 1, 0),
        ("TX04", "Phạm Minh Đức", "001090045678", "790045678", "D", "Long Biên, Hà Nội", date(1990, 1, 30), date(2019, 6, 20), 1, 1),
        ("TX05", "Vũ Văn Hùng", "001093056789", "790056789", "D", "Đống Đa, Hà Nội", date(1993, 9, 9), date(2021, 3, 5), 0, 1),
        ("TX06", "Đặng Văn Khoa", "001095067890", "790067890", "D", "Hai Bà Trưng, Hà Nội", date(1995, 4, 18), date(2022, 8, 1), 0, 1),
        ("TX07", "Ngô Thanh Long", "001096078901", "790078901", "D", "Nam Từ Liêm, Hà Nội", date(1996, 12, 5), date(2023, 2, 14), 0, 1),
    ]
    # (mã, số ngày trước, giờ, tuyến, xe, lái xe, phụ xe, số khách, giá vé); số ngày âm = chuyến sắp chạy
    plan = [
        ("C001", 55, 6, "T01", "29B-123.45", "TX01", "TX05", 40, 150000),
        ("C002", 52, 8, "T02", "29B-123.45", "TX02", "TX06", 43, 380000),
        ("C003", 50, 7, "T03", "29B-234.56", "TX03", "TX07", 30, 700000),
        ("C004", 44, 19, "T04", "30F-345.67", "TX04", "TX05", 25, 450000),
        ("C005", 40, 6, "T01", "29B-456.78", "TX01", "TX06", 12, 150000),
        ("C006", 35, 20, "T03", "29B-234.56", "TX03", "TX07", 32, 700000),
        ("C007", 30, 6, "T05", "30H-567.89", "TX02", "TX05", 41, 420000),
        ("C008", 22, 9, "T02", "29B-123.45", "TX04", "TX06", 42, 380000),
        ("C009", 18, 6, "T04", "30F-345.67", "TX01", "TX05", 27, 450000),
        ("C010", 12, 7, "T01", "29B-456.78", "TX02", "TX07", 14, 150000),
        ("C011", 9, 6, "T02", "29B-123.45", "TX01", "TX06", 40, 380000),
        ("C012", 6, 18, "T03", "30F-345.67", "TX03", "TX04", 27, 700000),
        ("C013", 3, 6, "T01", "29B-123.45", "TX02", "TX05", 38, 150000),
        ("C014", 1, 7, "T05", "30H-567.89", "TX04", "TX06", 40, 420000),
        ("C015", -3, 6, "T01", "29B-123.45", "TX01", "TX05", 30, 150000),
    ]
    chuyen = [(m, at(n, h), kh, gia, x, t, lai, phu) for m, n, h, t, x, lai, phu, kh, gia in plan]
    return {"MucPhucTap": muc, "Tuyen": tuyen, "Xe": xe, "TaiXe": taixe, "Chuyen": chuyen}
