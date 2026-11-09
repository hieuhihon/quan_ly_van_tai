"""Nhập dữ liệu từ file Excel, tạo file Excel mẫu và xuất bảng/báo cáo ra Excel."""
import re
from datetime import date, datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import database
import sample_data
from models import IMPORT_ORDER
from utils import AppError, fmt_value

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
KIND_TEXT = {"text": "Văn bản", "int": "Số nguyên", "float": "Số", "date": "Ngày (dd/mm/yyyy)",
             "datetime": "Ngày giờ (dd/mm/yyyy hh:mm)", "bool": "1 = có; 0 hoặc để trống = không"}
THIN = Side(style="thin", color="BBBBBB")


def _norm(s):
    return re.sub(r"[\s_\-]+", "", str(s).strip().lower())


# ------------------------------------------------------------------ nhập từ Excel
def _find_sheet(wb, ent):
    wanted = {_norm(ent.table), _norm(ent.title)}
    for name in wb.sheetnames:
        if _norm(name) in wanted:
            return wb[name]
    return None


def _map_headers(ent, header):
    names = {}
    for f in ent.fields:
        names[_norm(f.key)] = f.key
        names[_norm(f.label)] = f.key
    colmap = {}
    for idx, h in enumerate(header):
        key = names.get(_norm(h)) if h is not None else None
        if key and key not in colmap:
            colmap[key] = idx
    return colmap


def import_workbook(conn, path):
    """Nhập các sheet Tuyen, Xe, TaiXe, Chuyen, MucPhucTap. Khóa chính trùng thì cập nhật.

    Trả về {"summary": [(tên bảng, số dòng thêm, số dòng cập nhật, số dòng lỗi)], "errors": [thông báo]}.
    Dòng lỗi được bỏ qua và ghi lại, các dòng hợp lệ vẫn được nhập.
    """
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        raise AppError(f"Không đọc được file Excel: {e}")
    summary, errors, found = [], [], False
    try:
        for ent in IMPORT_ORDER:
            ws = _find_sheet(wb, ent)
            if ws is None:
                continue
            found = True
            rows = ws.iter_rows(values_only=True)
            colmap = _map_headers(ent, next(rows, None) or ())
            missing = [f.label for f in ent.fields if f.required and f.key not in colmap]
            if missing:
                errors.append(f"[{ent.table}] Thiếu cột bắt buộc: {', '.join(missing)}. Bỏ qua sheet này.")
                summary.append((ent.title, 0, 0, 0))
                continue
            added = updated = failed = 0
            for n, row in enumerate(rows, start=2):
                if row is None or all(c is None or str(c).strip() == "" for c in row):
                    continue
                raw = {k: (row[i] if i < len(row) else None) for k, i in colmap.items()}
                try:
                    status = database.upsert(conn, ent, raw)
                except AppError as e:
                    failed += 1
                    errors.append(f"[{ent.table}] dòng {n}: {e}")
                    continue
                if status == "them":
                    added += 1
                else:
                    updated += 1
            summary.append((ent.title, added, updated, failed))
    finally:
        wb.close()
    if not found:
        raise AppError("Không tìm thấy sheet hợp lệ. File cần có các sheet: Tuyen, Xe, TaiXe, Chuyen "
                       "(xem file mẫu).")
    return {"summary": summary, "errors": errors}


# ------------------------------------------------------------------ file Excel mẫu
def _style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def create_template(path, today=None):
    """Tạo file Excel mẫu gồm sheet hướng dẫn và các sheet dữ liệu (có sẵn dữ liệu minh họa)."""
    sample = sample_data.build(today)
    wb = Workbook()
    guide = wb.active
    guide.title = "HuongDan"
    lines = [
        ("HƯỚNG DẪN NHẬP LIỆU", True),
        ("1. Mỗi sheet là một bảng dữ liệu: MucPhucTap, Tuyen, Xe, TaiXe, Chuyen. Không đổi tên sheet và dòng tiêu đề (dòng 1).", False),
        ("2. Chỉ sửa/thêm dữ liệu từ dòng 2 trở xuống. Các dòng hiện có là DỮ LIỆU MẪU, hãy xóa hoặc thay bằng dữ liệu thật.", False),
        ("3. Nếu mã (khóa chính) đã có trong chương trình thì dòng đó sẽ CẬP NHẬT bản ghi cũ, chưa có thì THÊM MỚI.", False),
        ("4. Thứ tự nhập tự động: Tuyen -> Xe -> TaiXe -> Chuyen. Chuyến chỉ nhập được nếu xe, tuyến, lái xe, phụ xe đã tồn tại.", False),
        ("5. Cột LaiXe / PhuXe của sheet TaiXe: ghi 1 nếu tài xế đảm nhận vai trò đó (có thể cả hai), 0 hoặc để trống nếu không.", False),
        ("6. Ngày nhập dạng dd/mm/yyyy, ngày giờ dạng dd/mm/yyyy hh:mm (hoặc để Excel tự định dạng ngày).", False),
        ("7. Ràng buộc: lái xe khác phụ xe; số khách <= số ghế - 2; số khách >= 0; giá vé > 0; mức phức tạp thuộc {1, 2, 3}.", False),
        ("8. Dòng vi phạm ràng buộc sẽ bị bỏ qua và được liệt kê lý do sau khi nhập; các dòng hợp lệ vẫn được nhập.", False),
        ("9. Chuyến có ngày khởi hành trong tương lai (vd. C015) chưa tính vào lương, doanh thu và lịch bảo dưỡng.", False),
        ("", False),
    ]
    for i, (text, bold) in enumerate(lines, start=1):
        c = guide.cell(row=i, column=1, value=text)
        c.font = Font(name=FONT, bold=bold, size=14 if bold else 11)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        guide.merge_cells(start_row=i, start_column=1, end_row=i, end_column=5)
        guide.row_dimensions[i].height = 22 if bold else 18 * (len(text) // 105 + 1)
    r = len(lines) + 1
    for j, h in enumerate(("Sheet", "Tên cột", "Ý nghĩa", "Kiểu dữ liệu", "Bắt buộc"), start=1):
        c = guide.cell(row=r, column=j, value=h)
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
    for ent in IMPORT_ORDER:
        for f in ent.fields:
            r += 1
            values = (ent.table, f.key, f.label, KIND_TEXT[f.kind], "Có" if f.required else "Không")
            for j, v in enumerate(values, start=1):
                guide.cell(row=r, column=j, value=v).font = Font(name=FONT)
    for col, w in zip("ABCDE", (16, 18, 30, 40, 10)):
        guide.column_dimensions[col].width = w

    for ent in IMPORT_ORDER:
        ws = wb.create_sheet(ent.table)
        ws.append([f.key for f in ent.fields])
        _style_header(ws, len(ent.fields))
        for row in sample[ent.table]:
            ws.append(list(row))
        for j, f in enumerate(ent.fields, start=1):
            letter = get_column_letter(j)
            width = max(len(f.key), *(len(str(r[j - 1])) for r in sample[ent.table])) + 4
            ws.column_dimensions[letter].width = min(max(width, 12), 30)
            fmt = {"text": "@", "date": "DD/MM/YYYY", "datetime": "DD/MM/YYYY HH:MM"}.get(f.kind)
            for rr in range(2, 302):
                cell = ws.cell(row=rr, column=j)
                cell.font = Font(name=FONT)
                if fmt:
                    cell.number_format = fmt
    wb.save(path)


# ------------------------------------------------------------------ xuất ra Excel
def _excel_value(kind, v):
    if v is None:
        return None
    if kind == "date":
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    if kind == "datetime":
        return datetime.strptime(str(v)[:16], "%Y-%m-%d %H:%M")
    if kind == "bool":
        return "Có" if v else ""
    return v


def export_table(path, title, headers, rows, kinds):
    """Xuất bảng (giá trị thô + kiểu cột) ra file Excel, số/ngày được lưu đúng kiểu để người dùng tính toán tiếp."""
    wb = Workbook()
    ws = wb.active
    ws.title = re.sub(r"[\\/*?:\[\]]", "-", title)[:31] or "Du lieu"
    ws.append(list(headers))
    _style_header(ws, len(headers))
    for row in rows:
        ws.append([_excel_value(k, v) for k, v in zip(kinds, row)])
    formats = {"date": "DD/MM/YYYY", "datetime": "DD/MM/YYYY HH:MM", "float": "#,##0.##", "int": "0"}
    for j, kind in enumerate(kinds, start=1):
        width = max([len(str(headers[j - 1]))] + [len(fmt_value(kind, r[j - 1])) for r in rows]) + 3
        ws.column_dimensions[get_column_letter(j)].width = min(max(width, 10), 40)
        for rr in range(2, len(rows) + 2):
            cell = ws.cell(row=rr, column=j)
            cell.font = Font(name=FONT)
            cell.border = Border(bottom=THIN)
            if kind in formats:
                cell.number_format = formats[kind]
    wb.save(path)
