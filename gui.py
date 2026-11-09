"""Giao diện người dùng (Tkinter): các tab quản lý CRUD, báo cáo, nhập/xuất Excel."""
import os
import tempfile
import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import database
import excel_io
import sample_data
import services
from models import ENTITIES
from utils import AppError, edit_text, fmt_value, parse_date, sort_key, fmt_date

FONT = ("Segoe UI", 10)
BOLD = ("Segoe UI", 10, "bold")
HINTS = {"date": " (dd/mm/yyyy)", "datetime": " (dd/mm/yyyy hh:mm)"}
XLSX = [("Excel", "*.xlsx")]

GUIDE = """HƯỚNG DẪN NHANH

1. Nạp dữ liệu
   - Bấm "Tải file Excel mẫu", điền dữ liệu vào file rồi bấm "Nhập dữ liệu từ Excel".
   - Hoặc bấm "Nạp dữ liệu demo" để có ngay dữ liệu minh họa.

2. Thêm / sửa / xóa
   - Điền các ô ở khung "Thông tin" rồi bấm "Thêm mới".
   - Muốn sửa: bấm chọn một dòng trong bảng, sửa các ô rồi bấm "Cập nhật" (mã không sửa được).
   - Muốn xóa: chọn dòng rồi bấm "Xóa". Bản ghi đang được dùng ở bảng khác sẽ không xóa được.
   - Bấm tiêu đề cột để sắp xếp; gõ vào ô "Tìm kiếm" để lọc.

3. Thứ tự nhập tay: Mức phức tạp -> Tuyến xe -> Xe khách -> Tài xế -> Chuyến xe.
   Tài xế phải được tick "Là lái xe" / "Là phụ xe" thì mới chọn được trong Chuyến xe.

4. Báo cáo: tab "Báo cáo" gồm lương tài xế theo tháng, doanh thu từng xe theo khoảng ngày
   và danh sách xe quá hạn bảo dưỡng. Mỗi bảng đều có nút "Xuất Excel".

Ghi chú: ô có dấu * là bắt buộc. Dòng màu đỏ trong tab Xe khách là xe quá hạn bảo dưỡng."""


class DataTable(ttk.Frame):
    """Bảng (Treeview) có ô tìm kiếm, sắp xếp khi bấm tiêu đề và thanh cuộn."""

    def __init__(self, master):
        super().__init__(master)
        self.headers, self.kinds, self.rows, self.tags, self.order = [], [], [], [], []
        self.sort_rev = False
        self.query = tk.StringVar()
        self.count = tk.StringVar()
        self.top = ttk.Frame(self)
        self.top.pack(fill="x", pady=(0, 6))
        ttk.Label(self.top, text="Tìm kiếm:").pack(side="left")
        ttk.Entry(self.top, textvariable=self.query, width=30).pack(side="left", padx=6)
        ttk.Label(self.top, textvariable=self.count, foreground="#555").pack(side="left", padx=10)
        self.query.trace_add("write", lambda *_: self.apply())
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(body, show="headings", selectmode="browse", height=10)
        ys = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        body.grid_rowconfigure(0, weight=1)
        body.grid_columnconfigure(0, weight=1)
        self.tree.tag_configure("odd", background="#f3f7fb")
        self.tree.tag_configure("overdue", background="#fde7e9", foreground="#b00020")

    def load(self, headers, kinds, rows, tags=None):
        self.headers, self.kinds, self.rows = headers, kinds, rows
        self.tags = tags or [""] * len(rows)
        ids = [f"c{i}" for i in range(len(headers))]
        self.tree["columns"] = ids
        for i, (cid, h) in enumerate(zip(ids, headers)):
            self.tree.heading(cid, text=h, command=lambda c=i: self.sort_by(c))
            data_w = max([len(fmt_value(kinds[i], r[i])) for r in rows[:300]] + [0]) * 8 + 24
            self.tree.column(cid, width=min(max(len(h) * 10 + 28, data_w, 80), 320), stretch=False,
                             anchor="e" if kinds[i] in ("int", "float") else "w")
        self.order = list(range(len(rows)))
        self.apply()

    def apply(self):
        q = self.query.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        shown = 0
        for i in self.order:
            cells = [fmt_value(k, v) for k, v in zip(self.kinds, self.rows[i])]
            if q and not any(q in c.lower() for c in cells):
                continue
            tag = self.tags[i] or ("odd" if shown % 2 else "")
            self.tree.insert("", "end", iid=str(i), values=cells, tags=(tag,) if tag else ())
            shown += 1
        self.count.set(f"{shown}/{len(self.rows)} dòng")

    def sort_by(self, col):
        self.sort_rev = not self.sort_rev
        self.order.sort(key=lambda i: sort_key(self.rows[i][col]), reverse=self.sort_rev)
        self.apply()

    def selected_index(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None


class CrudTab(ttk.Frame):
    """Tab quản lý một bảng: danh sách + form Thêm / Cập nhật / Xóa."""

    def __init__(self, master, app, ent):
        super().__init__(master, padding=8)
        self.app, self.conn, self.ent = app, app.conn, ent
        self.table = DataTable(self)
        self.table.pack(fill="both", expand=True)
        ttk.Button(self.table.top, text="Xuất Excel", command=self.export).pack(side="right")
        self.table.tree.bind("<<TreeviewSelect>>", self.on_select)

        form = ttk.LabelFrame(self, text=f"Thông tin {ent.title.lower()}", padding=8)
        form.pack(fill="x", pady=(8, 0))
        self.vars, self.inputs, self.combos = {}, {}, {}
        for n, f in enumerate(ent.fields):
            r, c = divmod(n, 3)
            c *= 2
            if f.kind == "bool":
                var = tk.BooleanVar()
                w = ttk.Checkbutton(form, text=f.label, variable=var)
                w.grid(row=r, column=c, columnspan=2, sticky="w", padx=8, pady=4)
            else:
                text = f.label + HINTS.get(f.kind, "") + (" *" if f.required else "")
                ttk.Label(form, text=text, wraplength=170, justify="right").grid(row=r, column=c, sticky="e", padx=(8, 4), pady=4)
                var = tk.StringVar()
                if f.lookup:
                    w = ttk.Combobox(form, textvariable=var, state="readonly", width=28)
                    self.combos[f.key] = w
                else:
                    w = ttk.Entry(form, textvariable=var, width=24)
                w.grid(row=r, column=c + 1, sticky="w", pady=4)
            self.vars[f.key], self.inputs[f.key] = var, w
        btns = ttk.Frame(form)
        btns.grid(row=(len(ent.fields) - 1) // 3 + 1, column=0, columnspan=6, sticky="w", pady=(8, 0))
        ttk.Button(btns, text="Thêm mới", style="Accent.TButton", command=self.add).pack(side="left", padx=4)
        ttk.Button(btns, text="Cập nhật", command=self.update_row).pack(side="left", padx=4)
        ttk.Button(btns, text="Xóa", command=self.remove).pack(side="left", padx=4)
        ttk.Button(btns, text="Nhập mới (xóa trắng ô nhập)", command=self.clear_form).pack(side="left", padx=4)
        self.reload()

    # ---- dữ liệu
    def reload(self):
        for f in self.ent.fields:
            if f.lookup:
                self.combos[f.key]["values"] = database.lookup(self.conn, f.lookup)
        headers, kinds, rows, tags = services.table_data(self.conn, self.ent)
        self.table.load(headers, kinds, rows, tags)

    def on_show(self):
        self.reload()

    # ---- form
    def read_form(self):
        raw = {}
        for f in self.ent.fields:
            v = self.vars[f.key].get()
            raw[f.key] = v.split(" - ")[0].strip() if f.lookup and v else v
        return raw

    def fill_form(self, d):
        for f in self.ent.fields:
            v = d.get(f.key)
            if f.kind == "bool":
                self.vars[f.key].set(bool(v))
            elif f.lookup:
                opts = self.combos[f.key]["values"]
                self.vars[f.key].set(next((o for o in opts if o.split(" - ")[0] == str(v)), "" if v is None else str(v)))
            else:
                self.vars[f.key].set(edit_text(f.kind, v))
        self.inputs[self.ent.pk].configure(state="disabled")

    def clear_form(self):
        for f in self.ent.fields:
            self.vars[f.key].set(False if f.kind == "bool" else "")
        self.inputs[self.ent.pk].configure(state="normal")
        self.table.tree.selection_set(())

    def on_select(self, _event=None):
        i = self.table.selected_index()
        if i is None:
            return
        d = database.get(self.conn, self.ent, self.table.rows[i][0])
        if d:
            self.fill_form(d)

    # ---- thao tác
    def add(self):
        try:
            database.insert(self.conn, self.ent, self.read_form())
        except AppError as e:
            messagebox.showerror("Không thể thêm", str(e))
            return
        self.app.status.set(f"Đã thêm {self.ent.title.lower()} mới.")
        self.clear_form()
        self.app.refresh_all()

    def update_row(self):
        i = self.table.selected_index()
        if i is None:
            messagebox.showinfo("Cập nhật", "Hãy bấm chọn một dòng trong bảng trước.")
            return
        raw = self.read_form()
        raw[self.ent.pk] = self.table.rows[i][0]
        try:
            database.update(self.conn, self.ent, raw)
        except AppError as e:
            messagebox.showerror("Không thể cập nhật", str(e))
            return
        self.app.status.set(f"Đã cập nhật {self.ent.title.lower()}.")
        self.clear_form()
        self.app.refresh_all()

    def remove(self):
        i = self.table.selected_index()
        if i is None:
            messagebox.showinfo("Xóa", "Hãy bấm chọn một dòng trong bảng trước.")
            return
        pk = self.table.rows[i][0]
        if not messagebox.askyesno("Xác nhận xóa", f"Xóa {self.ent.title.lower()} '{pk}'?"):
            return
        try:
            database.delete(self.conn, self.ent, pk)
        except AppError as e:
            messagebox.showerror("Không thể xóa", str(e))
            return
        self.app.status.set(f"Đã xóa {self.ent.title.lower()} '{pk}'.")
        self.clear_form()
        self.app.refresh_all()

    def export(self):
        self.app.export_table(self.ent.title, self.table)


class ReportTab(ttk.Frame):
    """Khung chung cho các báo cáo: vùng chọn tham số + bảng kết quả + xuất Excel."""
    title = "Báo cáo"

    def __init__(self, master, app):
        super().__init__(master, padding=8)
        self.app, self.conn = app, app.conn
        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 6))
        self.controls(top)
        ttk.Button(top, text="Xem báo cáo", style="Accent.TButton", command=self.run).pack(side="left", padx=10)
        ttk.Button(top, text="Xuất Excel", command=lambda: app.export_table(self.title, self.table)).pack(side="left")
        self.table = DataTable(self)
        self.table.pack(fill="both", expand=True)
        self.summary = tk.StringVar()
        ttk.Label(self, textvariable=self.summary, font=BOLD).pack(anchor="w", pady=(6, 0))

    def controls(self, parent):
        pass

    def fetch(self):
        raise NotImplementedError

    def on_show(self):
        self.run()

    def run(self):
        try:
            headers, kinds, rows, tags, summary = self.fetch()
        except AppError as e:
            messagebox.showwarning(self.title, str(e))
            return
        self.table.load(headers, kinds, rows, tags)
        self.summary.set(summary)


class SalaryTab(ReportTab):
    title = "Lương tài xế"

    def controls(self, parent):
        today = date.today()
        self.year, self.month = tk.StringVar(value=str(today.year)), tk.StringVar(value=str(today.month))
        ttk.Label(parent, text="Tháng:").pack(side="left")
        ttk.Spinbox(parent, from_=1, to=12, width=4, textvariable=self.month).pack(side="left", padx=(4, 12))
        ttk.Label(parent, text="Năm:").pack(side="left")
        ttk.Spinbox(parent, from_=2000, to=2100, width=6, textvariable=self.year).pack(side="left", padx=4)

    def fetch(self):
        try:
            y, m = int(self.year.get()), int(self.month.get())
        except ValueError:
            raise AppError("Tháng và năm phải là số nguyên.")
        rows = services.salary_report(self.conn, y, m)
        headers = ["Mã tài xế", "Họ tên", "Số chuyến lái xe", "Số chuyến phụ xe", "Lương tháng (VND)"]
        total = sum(r[4] for r in rows)
        return (headers, ["text", "text", "int", "int", "float"], rows, None,
                f"Tổng lương tháng {m:02d}/{y} (tính đến hiện tại): {total:,.0f} VND")


class RevenueTab(ReportTab):
    title = "Doanh thu xe"

    def controls(self, parent):
        today = date.today()
        self.d1 = tk.StringVar(value=fmt_date((today - timedelta(days=90)).isoformat()))
        self.d2 = tk.StringVar(value=fmt_date(today.isoformat()))
        ttk.Label(parent, text="Từ ngày (dd/mm/yyyy):").pack(side="left")
        ttk.Entry(parent, textvariable=self.d1, width=12).pack(side="left", padx=(4, 12))
        ttk.Label(parent, text="Đến ngày:").pack(side="left")
        ttk.Entry(parent, textvariable=self.d2, width=12).pack(side="left", padx=4)

    def fetch(self):
        try:
            d1, d2 = parse_date(self.d1.get()), parse_date(self.d2.get())
        except ValueError:
            raise AppError("Ngày không hợp lệ, hãy nhập theo dạng dd/mm/yyyy.")
        rows = services.revenue_report(self.conn, d1, d2)
        headers = ["Biển số", "Xe", "Số chuyến", "Tổng số khách", "Doanh thu (VND)"]
        total = sum(r[4] for r in rows)
        return (headers, ["text", "text", "int", "int", "float"], rows, None,
                f"Tổng doanh thu từ {fmt_date(d1)} đến {fmt_date(d2)}: {total:,.0f} VND")


class MaintenanceTab(ReportTab):
    title = "Xe quá hạn bảo dưỡng"

    def controls(self, parent):
        self.only_overdue = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent, text="Chỉ hiện xe quá hạn bảo dưỡng", variable=self.only_overdue,
                        command=self.run).pack(side="left")

    def fetch(self):
        data = services.maintenance(self.conn)
        shown = [m for m in data if m["QuaHan"]] if self.only_overdue.get() else data
        headers = ["Biển số", "Xe", "Ngày BD cuối", "Số km quy đổi", "Số ngày bị giảm trừ",
                   "Ngày BD kế tiếp", "Số ngày còn lại", "Tình trạng"]
        kinds = ["text", "text", "date", "float", "int", "date", "int", "text"]
        rows = [[m["BienSo"], m["TenXe"], m["NgayBDCuoi"], m["KmQuyDoi"], m["NgayGiam"], m["NgayBDKeTiep"],
                 m["SoNgayConLai"], "Quá hạn" if m["QuaHan"] else "Còn hạn"] for m in shown]
        tags = ["overdue" if m["QuaHan"] else "" for m in shown]
        over = sum(1 for m in data if m["QuaHan"])
        return headers, kinds, rows, tags, f"Có {over}/{len(data)} xe quá hạn bảo dưỡng."


class ReportsTab(ttk.Frame):
    def __init__(self, master, app):
        super().__init__(master, padding=4)
        self.inner = ttk.Notebook(self)
        self.inner.pack(fill="both", expand=True)
        for cls in (SalaryTab, RevenueTab, MaintenanceTab):
            self.inner.add(cls(self.inner, app), text=cls.title)
        self.inner.bind("<<NotebookTabChanged>>", lambda _e: self.on_show())

    def on_show(self):
        self.nametowidget(self.inner.select()).on_show()


class App(tk.Tk):
    def __init__(self, db_path=database.DB_PATH):
        super().__init__()
        self.title("Quản lý hãng vận tải hành khách đường dài")
        self.geometry("1260x780")
        self.minsize(1020, 640)
        self.conn = database.connect(db_path)
        self.status = tk.StringVar(value="Sẵn sàng.")
        self._style()
        self._menu()

        bar = ttk.Frame(self, padding=(8, 8, 8, 0))
        bar.pack(fill="x")
        for text, cmd in (("Nhập dữ liệu từ Excel", self.import_excel), ("Tải file Excel mẫu", self.save_template),
                          ("Nạp dữ liệu demo", self.load_demo)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=4)
        ttk.Button(bar, text="Hướng dẫn", command=self.show_guide).pack(side="right", padx=4)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tabs = []
        for ent in ENTITIES:
            tab = CrudTab(self.nb, self, ent)
            self.nb.add(tab, text=ent.title)
            self.tabs.append(tab)
        self.reports = ReportsTab(self.nb, self)
        self.nb.add(self.reports, text="Báo cáo")
        self.nb.bind("<<NotebookTabChanged>>", lambda _e: self.refresh_current())
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w", padding=(8, 3)).pack(fill="x", side="bottom")
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _style(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure(".", font=FONT)
        s.configure("Treeview", rowheight=26, font=FONT)
        s.configure("Treeview.Heading", font=BOLD, padding=4)
        s.configure("TNotebook.Tab", padding=(16, 7), font=BOLD)
        s.configure("Accent.TButton", background="#1565c0", foreground="white", font=BOLD, padding=(12, 4))
        s.map("Accent.TButton", background=[("active", "#0d47a1")])

    def _menu(self):
        m = tk.Menu(self)
        f = tk.Menu(m, tearoff=0)
        f.add_command(label="Nhập dữ liệu từ Excel...", command=self.import_excel)
        f.add_command(label="Tải file Excel mẫu...", command=self.save_template)
        f.add_command(label="Nạp dữ liệu demo", command=self.load_demo)
        f.add_separator()
        f.add_command(label="Xóa toàn bộ dữ liệu...", command=self.clear_data)
        f.add_separator()
        f.add_command(label="Thoát", command=self.on_close)
        m.add_cascade(label="Tệp", menu=f)
        h = tk.Menu(m, tearoff=0)
        h.add_command(label="Hướng dẫn nhanh", command=self.show_guide)
        m.add_cascade(label="Trợ giúp", menu=h)
        self.config(menu=m)

    # ---- làm mới
    def refresh_all(self):
        for t in self.tabs:
            t.reload()
        self.refresh_current()

    def refresh_current(self):
        self.nametowidget(self.nb.select()).on_show()

    # ---- Excel
    def import_excel(self):
        path = filedialog.askopenfilename(title="Chọn file Excel dữ liệu", filetypes=XLSX)
        if path:
            self._import(path)

    def _import(self, path):
        try:
            res = excel_io.import_workbook(self.conn, path)
        except AppError as e:
            messagebox.showerror("Nhập Excel", str(e))
            return
        lines = [f"{t}: thêm {a}, cập nhật {u}, lỗi {e}" for t, a, u, e in res["summary"]]
        if res["errors"]:
            lines += ["", "CÁC DÒNG BỊ BỎ QUA:"] + res["errors"]
        self.status.set("Đã nhập xong file Excel.")
        self.refresh_all()
        self.show_text("Kết quả nhập Excel", "\n".join(lines))

    def save_template(self):
        path = filedialog.asksaveasfilename(title="Lưu file Excel mẫu", defaultextension=".xlsx",
                                            initialfile="mau_nhap_lieu.xlsx", filetypes=XLSX)
        if path:
            excel_io.create_template(path)
            self.status.set(f"Đã lưu file mẫu: {path}")
            messagebox.showinfo("File mẫu", f"Đã lưu file mẫu tại:\n{path}\n\nXem sheet HuongDan để biết cách điền.")

    def load_demo(self):
        if not messagebox.askyesno("Dữ liệu demo", "Thêm dữ liệu minh họa vào chương trình?\n"
                                   "(Bản ghi trùng mã sẽ được cập nhật, dữ liệu khác được giữ nguyên.)"):
            return
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "demo.xlsx")
            excel_io.create_template(path)
            self._import(path)

    def export_table(self, title, table):
        if not table.rows:
            messagebox.showinfo("Xuất Excel", "Không có dữ liệu để xuất.")
            return
        path = filedialog.asksaveasfilename(title="Xuất ra Excel", defaultextension=".xlsx",
                                            initialfile=f"{title}.xlsx".replace(" ", "_"), filetypes=XLSX)
        if path:
            excel_io.export_table(path, title, table.headers, table.rows, table.kinds)
            self.status.set(f"Đã xuất {len(table.rows)} dòng ra {path}")

    def clear_data(self):
        if messagebox.askyesno("Xóa toàn bộ dữ liệu", "Xóa TẤT CẢ dữ liệu trong chương trình?\nThao tác này không hoàn tác được.",
                               icon="warning"):
            database.clear_all(self.conn)
            self.refresh_all()
            self.status.set("Đã xóa toàn bộ dữ liệu.")

    # ---- hộp thoại
    def show_text(self, title, text):
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("760x460")
        box = ScrolledText(win, font=("Consolas", 10), wrap="word")
        box.pack(fill="both", expand=True, padx=8, pady=8)
        box.insert("1.0", text)
        box.configure(state="disabled")
        ttk.Button(win, text="Đóng", command=win.destroy).pack(pady=(0, 8))

    def show_guide(self):
        self.show_text("Hướng dẫn nhanh", GUIDE)

    def on_close(self):
        self.conn.close()
        self.destroy()
