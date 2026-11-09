"""Mô tả các thực thể (bảng) của hệ thống. GUI, nhập Excel và CSDL cùng dùng chung mô tả này."""
from dataclasses import dataclass


@dataclass(frozen=True)
class Field:
    key: str                  # tên cột trong CSDL / tiêu đề cột trong Excel
    label: str                # nhãn hiển thị cho người dùng
    kind: str = "text"        # text | int | float | date | datetime | bool
    required: bool = True
    lookup: str = ""          # tên danh mục cho ô chọn (combobox): xe | tuyen | laixe | phuxe | muc
    virtual: bool = False     # không phải cột thật của bảng (vai trò lái xe / phụ xe)


@dataclass(frozen=True)
class Entity:
    table: str
    title: str
    pk: str
    fields: tuple
    derived: tuple = ()       # cột dẫn xuất chỉ đọc: (key, nhãn, kind)


MUC = Entity("MucPhucTap", "Mức phức tạp", "Muc", (
    Field("Muc", "Mức (1, 2, 3)", "int"),
    Field("HeSoLuong", "Hệ số lương", "float"),
    Field("HeSoKmDuongKho", "Hệ số km đường khó", "float"),
))

XE = Entity("Xe", "Xe khách", "BienSo", (
    Field("BienSo", "Biển số"),
    Field("MauXe", "Màu xe"),
    Field("HangSX", "Hãng sản xuất"),
    Field("DoiXe", "Đời xe", "int"),
    Field("Model", "Model"),
    Field("SoGhe", "Số ghế", "int"),
    Field("NgayBDCuoi", "Ngày bảo dưỡng cuối", "date"),
), derived=(
    ("SoNamSuDung", "Số năm sử dụng", "int"),
    ("NgayBDKeTiep", "Ngày BD kế tiếp", "date"),
    ("SoNgayConLai", "Số ngày còn lại", "int"),
))

TAIXE = Entity("TaiXe", "Tài xế", "MaTX", (
    Field("MaTX", "Mã tài xế"),
    Field("HoTen", "Họ tên"),
    Field("CMT", "Chứng minh thư"),
    Field("SoBangLai", "Số bằng lái"),
    Field("LoaiBang", "Loại bằng lái"),
    Field("DiaChi", "Địa chỉ", required=False),
    Field("NgaySinh", "Ngày sinh", "date"),
    Field("NgayVaoLam", "Ngày vào làm", "date"),
    Field("LaiXe", "Là lái xe", "bool", required=False, virtual=True),
    Field("PhuXe", "Là phụ xe", "bool", required=False, virtual=True),
), derived=(
    ("ThamNien", "Thâm niên (năm)", "int"),
))

TUYEN = Entity("Tuyen", "Tuyến xe", "MaTuyen", (
    Field("MaTuyen", "Mã tuyến"),
    Field("DiemDau", "Điểm đầu"),
    Field("DiemCuoi", "Điểm cuối"),
    Field("DoDaiKm", "Độ dài (km)", "float"),
    Field("MucPhucTap", "Mức phức tạp", "int", lookup="muc"),
))

CHUYEN = Entity("Chuyen", "Chuyến xe", "MaChuyen", (
    Field("MaChuyen", "Mã chuyến"),
    Field("NgayKhoiHanh", "Ngày khởi hành", "datetime"),
    Field("SoKhach", "Số khách", "int"),
    Field("GiaVe", "Giá vé (VND)", "float"),
    Field("BienSo", "Xe (biển số)", lookup="xe"),
    Field("MaTuyen", "Tuyến", lookup="tuyen"),
    Field("MaLaiXe", "Lái xe", lookup="laixe"),
    Field("MaPhuXe", "Phụ xe", lookup="phuxe"),
), derived=(
    ("DoanhThu", "Doanh thu (VND)", "float"),
))

ENTITIES = (XE, TAIXE, TUYEN, CHUYEN, MUC)          # thứ tự hiển thị các tab
IMPORT_ORDER = (MUC, TUYEN, XE, TAIXE, CHUYEN)      # thứ tự nhập Excel (bảng cha trước bảng con)
