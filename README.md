# Quản lý hãng vận tải hành khách đường dài

Bài tập lớn Cơ sở dữ liệu – Project 5. Ứng dụng desktop viết bằng **Python**, dùng **SQLite** làm CSDL,
giao diện **Tkinter** (có sẵn trong Python) và **openpyxl** để nhập/xuất Excel.
Người dùng không cần biết lập trình: chỉ cần mở chương trình, bấm nút.

## Chức năng

- **CRUD đầy đủ** (thêm / xem / sửa / xóa) cho: Xe khách, Tài xế, Tuyến xe, Chuyến xe, Mức phức tạp.
- **Nhập dữ liệu từ file Excel** (có file mẫu và sheet hướng dẫn); dòng sai được bỏ qua và báo rõ lý do, dòng đúng vẫn được nhập.
- **Xuất Excel** từ mọi bảng và mọi báo cáo.
- **Kiểm tra ràng buộc** ngay trong CSDL (CHECK, UNIQUE, FOREIGN KEY, TRIGGER) và báo lỗi tiếng Việt:
  R1 lái xe ≠ phụ xe · R2 số khách ≤ số ghế − 2 · R3 mức phức tạp ∈ {1, 2, 3} · R4 mỗi chuyến đúng 1 lái + 1 phụ · R5 số khách ≥ 0, giá vé > 0.
- **Thuộc tính dẫn xuất** hiển thị tự động: số năm sử dụng, ngày bảo dưỡng kế tiếp, số ngày còn lại, thâm niên, doanh thu chuyến.
- **Báo cáo**: lương tài xế theo tháng · doanh thu từng xe theo khoảng ngày · danh sách xe quá hạn bảo dưỡng.

## Cài đặt và chạy

Tải và giải nén file VanTaiHanhKhach.rar để sử dụng chương trìnhh

## Hướng dẫn sử dụng

1. **Thêm mới**: điền các ô ở khung "Thông tin" (ô có dấu `*` là bắt buộc) rồi bấm *Thêm mới*.
2. **Sửa**: bấm chọn một dòng trong bảng, sửa các ô rồi bấm *Cập nhật* (mã/khóa chính không sửa được).
3. **Xóa**: chọn dòng rồi bấm *Xóa*. Bản ghi đang được dùng (ví dụ xe đã có chuyến) sẽ không xóa được.
4. Bấm tiêu đề cột để sắp xếp; gõ vào ô *Tìm kiếm* để lọc. Dòng đỏ ở tab Xe khách là xe quá hạn bảo dưỡng.
5. Thứ tự nhập tay: Mức phức tạp → Tuyến xe → Xe khách → Tài xế → Chuyến xe.
   Tài xế cần tick *Là lái xe* / *Là phụ xe* thì mới được chọn trong Chuyến xe (chuyên biệt hóa chồng lẫn, bán phần).

### Định dạng file Excel nhập liệu

File gồm các sheet `MucPhucTap`, `Tuyen`, `Xe`, `TaiXe`, `Chuyen` (sheet nào không có thì bỏ qua). Dòng 1 là tiêu đề cột (đúng tên
thuộc tính trong tài liệu), dữ liệu từ dòng 2. Mã trùng với dữ liệu đang có thì **cập nhật**, chưa có thì **thêm mới**.
Chi tiết từng cột xem sheet `HuongDan` trong `data/mau_nhap_lieu.xlsx`. Ngày nhập dạng `dd/mm/yyyy`, ngày giờ `dd/mm/yyyy hh:mm`;
cột `LaiXe`/`PhuXe` của sheet `TaiXe` ghi `1` nếu có vai trò đó.

## Cấu trúc thư mục

```
main.py            điểm khởi chạy
gui.py             giao diện Tkinter (các tab CRUD, báo cáo, menu Excel)
database.py        kết nối SQLite, kiểm tra dữ liệu, CRUD tổng quát, việc đổi lỗi SQL thành tiếng Việt
services.py        nghiệp vụ: lương, doanh thu, lịch bảo dưỡng, thuộc tính dẫn xuất
excel_io.py        nhập Excel, tạo file mẫu, xuất Excel
models.py          mô tả thực thể/thuộc tính (dùng chung cho GUI, Excel, CSDL)
sample_data.py     dữ liệu mẫu cho file Excel mẫu và chế độ demo
schema.sql         lược đồ CSDL (bảng, khóa, ràng buộc, trigger)
data/              mau_nhap_lieu.xlsx (file mẫu); vantai.db được tạo khi chạy
tests/             kiểm thử tự động
```

## Lược đồ CSDL

| Bảng | Khóa chính | Khóa ngoại / ghi chú |
|---|---|---|
| `MucPhucTap` | `Muc` (1, 2, 3) | `HeSoLuong`, `HeSoKmDuongKho` |
| `Xe` | `BienSo` | |
| `TaiXe` | `MaTX` | `CMT` UNIQUE |
| `LaiXe`, `PhuXe` | `MaTX` | → `TaiXe` (chuyên biệt hóa chồng lẫn, bán phần) |
| `Tuyen` | `MaTuyen` | `MucPhucTap` → `MucPhucTap` (N:1) |
| `Chuyen` | `MaChuyen` | `BienSo` → `Xe`, `MaTuyen` → `Tuyen`, `MaLaiXe` → `LaiXe`, `MaPhuXe` → `PhuXe` (1:N) |

Ràng buộc R2 (liên bảng) cài bằng trigger; R1, R3, R5 bằng CHECK. Thuộc tính dẫn xuất không lưu trong CSDL mà tính khi truy vấn.

## Công thức nghiệp vụ

Hằng số nằm ở đầu `services.py` nên dễ chỉnh.

- **Lương mỗi chuyến** = 200.000 × hệ số độ phức tạp (cấp 1: 1.0, cấp 2: 1.2, cấp 3: 1.5) × hệ số vai trò (lái 2.0, phụ 1.0).
  Lương tháng = tổng các chuyến trong tháng đã khởi hành đến thời điểm hiện tại.
- **Bảo dưỡng**: km quy đổi = Σ (độ dài tuyến × hệ số đường khó: 1.0 / 1.5 / 2.0) của các chuyến đã chạy kể từ ngày bảo dưỡng cuối;
  ngày bị giảm = ⌊km quy đổi / 100⌋; số ngày còn lại = 360 − ngày giảm − số ngày đã trôi qua; ≤ 0 thì xe quá hạn.
  Ngày bảo dưỡng kế tiếp = ngày bảo dưỡng cuối + (360 − ngày giảm).
- **Doanh thu xe** = Σ (số khách × giá vé) của các chuyến đã thực hiện trong khoảng ngày chọn.

Giả định: "chuyến đã thực hiện" là chuyến có ngày khởi hành ≤ thời điểm hiện tại; số năm sử dụng = năm hiện tại − đời xe;
thâm niên = năm hiện tại − năm vào làm (theo `DATEDIFF(YEAR, …)`).




