-- Hệ thống quản lý hãng vận tải hành khách đường dài (SQLite)
-- Mô hình EER: Xe, Tuyến, Chuyến, Tài xế (chuyên biệt hóa chồng lẫn, bán phần: Lái xe / Phụ xe),
-- bảng tra cứu Mức phức tạp.
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS MucPhucTap (
    Muc             INTEGER PRIMARY KEY,
    HeSoLuong       REAL NOT NULL,          -- hệ số độ phức tạp dùng tính lương
    HeSoKmDuongKho  REAL NOT NULL,          -- hệ số km đường khó dùng tính lịch bảo dưỡng
    CONSTRAINT ck_muc  CHECK (Muc IN (1, 2, 3)),
    CONSTRAINT ck_heso CHECK (HeSoLuong > 0 AND HeSoKmDuongKho > 0)
);

CREATE TABLE IF NOT EXISTS Xe (
    BienSo      TEXT PRIMARY KEY,
    MauXe       TEXT NOT NULL,
    HangSX      TEXT NOT NULL,
    DoiXe       INTEGER NOT NULL,
    Model       TEXT NOT NULL,
    SoGhe       INTEGER NOT NULL,
    NgayBDCuoi  TEXT NOT NULL,              -- ISO: YYYY-MM-DD
    CONSTRAINT ck_soghe CHECK (SoGhe > 0),
    CONSTRAINT ck_doixe CHECK (DoiXe > 1900)
);

CREATE TABLE IF NOT EXISTS TaiXe (
    MaTX        TEXT PRIMARY KEY,
    HoTen       TEXT NOT NULL,
    CMT         TEXT NOT NULL UNIQUE,
    SoBangLai   TEXT NOT NULL,
    LoaiBang    TEXT NOT NULL,
    DiaChi      TEXT,
    NgaySinh    TEXT NOT NULL,
    NgayVaoLam  TEXT NOT NULL
);

-- Chuyên biệt hóa chồng lẫn (overlapping), bán phần (partial)
CREATE TABLE IF NOT EXISTS LaiXe (
    MaTX TEXT PRIMARY KEY REFERENCES TaiXe(MaTX) ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS PhuXe (
    MaTX TEXT PRIMARY KEY REFERENCES TaiXe(MaTX) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS Tuyen (
    MaTuyen     TEXT PRIMARY KEY,
    DiemDau     TEXT NOT NULL,
    DiemCuoi    TEXT NOT NULL,
    DoDaiKm     REAL NOT NULL,
    MucPhucTap  INTEGER NOT NULL REFERENCES MucPhucTap(Muc),
    CONSTRAINT ck_dodai     CHECK (DoDaiKm > 0),
    CONSTRAINT ck_muc_tuyen CHECK (MucPhucTap IN (1, 2, 3))
);

CREATE TABLE IF NOT EXISTS Chuyen (
    MaChuyen      TEXT PRIMARY KEY,
    NgayKhoiHanh  TEXT NOT NULL,            -- ISO: YYYY-MM-DD HH:MM
    SoKhach       INTEGER NOT NULL,
    GiaVe         REAL NOT NULL,
    BienSo        TEXT NOT NULL REFERENCES Xe(BienSo),
    MaTuyen       TEXT NOT NULL REFERENCES Tuyen(MaTuyen),
    MaLaiXe       TEXT NOT NULL REFERENCES LaiXe(MaTX),
    MaPhuXe       TEXT NOT NULL REFERENCES PhuXe(MaTX),
    CONSTRAINT ck_khach    CHECK (SoKhach >= 0),                 -- R5
    CONSTRAINT ck_gia      CHECK (GiaVe > 0),                    -- R5
    CONSTRAINT ck_lai_phu  CHECK (MaLaiXe <> MaPhuXe)            -- R1
);

CREATE INDEX IF NOT EXISTS idx_chuyen_xe   ON Chuyen(BienSo);
CREATE INDEX IF NOT EXISTS idx_chuyen_ngay ON Chuyen(NgayKhoiHanh);
CREATE INDEX IF NOT EXISTS idx_chuyen_lai  ON Chuyen(MaLaiXe);
CREATE INDEX IF NOT EXISTS idx_chuyen_phu  ON Chuyen(MaPhuXe);

-- R2: Số khách <= Số ghế - 2 (ràng buộc liên bảng nên dùng trigger)
CREATE TRIGGER IF NOT EXISTS trg_chuyen_suc_chua_ins
BEFORE INSERT ON Chuyen
WHEN NEW.SoKhach > (SELECT SoGhe FROM Xe WHERE BienSo = NEW.BienSo) - 2
BEGIN
    SELECT RAISE(ABORT, 'R2: Số khách phải nhỏ hơn hoặc bằng số ghế trừ 2.');
END;

CREATE TRIGGER IF NOT EXISTS trg_chuyen_suc_chua_upd
BEFORE UPDATE ON Chuyen
WHEN NEW.SoKhach > (SELECT SoGhe FROM Xe WHERE BienSo = NEW.BienSo) - 2
BEGIN
    SELECT RAISE(ABORT, 'R2: Số khách phải nhỏ hơn hoặc bằng số ghế trừ 2.');
END;

CREATE TRIGGER IF NOT EXISTS trg_xe_soghe_upd
BEFORE UPDATE OF SoGhe ON Xe
WHEN EXISTS (SELECT 1 FROM Chuyen WHERE BienSo = NEW.BienSo AND SoKhach > NEW.SoGhe - 2)
BEGIN
    SELECT RAISE(ABORT, 'R2: Đã có chuyến chở nhiều khách hơn số ghế mới trừ 2.');
END;

INSERT OR IGNORE INTO MucPhucTap (Muc, HeSoLuong, HeSoKmDuongKho) VALUES
    (1, 1.0, 1.0),
    (2, 1.2, 1.5),
    (3, 1.5, 2.0);
