"""Lưu / đọc KHĐT và lớp học trong SQLite (cùng file CSDL với bảng nhan_vien)."""
from datetime import datetime

from khdt_core import sap_xep

_COT_LOP = ("mien", "loai_hinh", "loai_hinh_nhan", "stt_goc", "thu_tu_goc", "ten_lop", "hinh_thuc",
            "tu_ngay", "den_ngay", "thoi_gian", "dia_diem", "giao_vien", "so_hv", "doi_tuong",
            "ghi_chu", "muc", "doi_tac", "dong_goc")


def khoi_tao_bang(conn):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS khdt (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            thang INTEGER NOT NULL, nam INTEGER NOT NULL, mien TEXT NOT NULL,
            sheet TEXT, ten_file TEXT, ngay_nhap TEXT,
            UNIQUE (thang, nam, mien)
        );
        CREATE TABLE IF NOT EXISTS lop_hoc (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            khdt_id INTEGER NOT NULL REFERENCES khdt(id) ON DELETE CASCADE,
            mien TEXT, loai_hinh TEXT, loai_hinh_nhan TEXT, stt_goc TEXT, thu_tu_goc INTEGER,
            ten_lop TEXT, hinh_thuc TEXT, tu_ngay TEXT, den_ngay TEXT, thoi_gian TEXT,
            dia_diem TEXT, giao_vien TEXT, so_hv INTEGER, doi_tuong TEXT, ghi_chu TEXT,
            muc TEXT, doi_tac TEXT, dong_goc INTEGER
        );
        CREATE TABLE IF NOT EXISTS lop_thang (
            lop_id INTEGER NOT NULL REFERENCES lop_hoc(id) ON DELETE CASCADE,
            thang TEXT NOT NULL  -- 'YYYY-MM'
        );
        CREATE INDEX IF NOT EXISTS idx_lop_khdt ON lop_hoc (khdt_id);
        CREATE INDEX IF NOT EXISTS idx_lop_thang ON lop_thang (thang, lop_id);
    """)
    # CSDL tạo từ phiên bản trước chưa có cột doi_tac
    if "doi_tac" not in {r[1] for r in conn.execute("PRAGMA table_info(lop_hoc)")}:
        conn.execute("ALTER TABLE lop_hoc ADD COLUMN doi_tac TEXT")
    conn.commit()


def luu_khdt(conn, thang, nam, mien, sheet, ten_file, lops):
    """Ghi đè KHĐT cùng (tháng, năm, miền). Trả về khdt_id."""
    conn.execute("PRAGMA foreign_keys = ON")
    with conn:  # transaction: lỗi giữa chừng thì không mất dữ liệu cũ
        conn.execute("DELETE FROM khdt WHERE thang = ? AND nam = ? AND mien = ?", (thang, nam, mien))
        cur = conn.execute(
            "INSERT INTO khdt (thang, nam, mien, sheet, ten_file, ngay_nhap) VALUES (?,?,?,?,?,?)",
            (thang, nam, mien, sheet, ten_file, datetime.now().isoformat(timespec="seconds")))
        khdt_id = cur.lastrowid
        for lop in lops:
            c = conn.execute(
                f"INSERT INTO lop_hoc (khdt_id, {', '.join(_COT_LOP)}) "
                f"VALUES (?, {', '.join('?' * len(_COT_LOP))})",
                (khdt_id, *[lop.get(k) for k in _COT_LOP]))
            conn.executemany("INSERT INTO lop_thang (lop_id, thang) VALUES (?, ?)",
                             [(c.lastrowid, t) for t in lop.get("thang", [])])
    return khdt_id


def ds_khdt(conn):
    return [dict(r) for r in conn.execute(
        "SELECT k.*, (SELECT COUNT(*) FROM lop_hoc l WHERE l.khdt_id = k.id) AS so_lop "
        "FROM khdt k ORDER BY nam DESC, thang DESC, mien")]


def doc_lop(conn, mien, nam, thang, ma_loai=None, che_do="khdt"):
    """Các lớp của miền trong tháng (nam, thang).

    Gồm lớp của KHĐT tháng đó + lớp kéo dài từ KHĐT tháng trước chạm sang tháng này
    (lớp tính vào tháng bắt đầu và cả các tháng sau). Loại trùng khi KHĐT tháng này
    đã ghi lại chính lớp đó.
    """
    key_thang = f"{nam:04d}-{thang:02d}"
    rows = conn.execute(
        "SELECT l.*, k.thang AS k_thang, k.nam AS k_nam FROM lop_hoc l "
        "JOIN khdt k ON k.id = l.khdt_id "
        "WHERE l.mien = ? AND l.id IN (SELECT lop_id FROM lop_thang WHERE thang = ?)",
        (mien, key_thang)).fetchall()
    lops, da_co = [], set()
    for r in rows:  # lớp thuộc chính KHĐT của tháng trước
        if (r["k_nam"], r["k_thang"]) == (nam, thang):
            d = dict(r)
            lops.append(d)
            da_co.add((d["ten_lop"], d["tu_ngay"], d["den_ngay"]))
    for r in rows:
        if (r["k_nam"], r["k_thang"]) != (nam, thang):
            d = dict(r)
            if (d["ten_lop"], d["tu_ngay"], d["den_ngay"]) not in da_co:
                d["keo_dai"] = True
                lops.append(d)
    if ma_loai:
        lops = [x for x in lops if x["loai_hinh"] in ma_loai]
    return sap_xep(lops, che_do)
