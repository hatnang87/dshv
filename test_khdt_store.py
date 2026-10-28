"""Kiểm thử nhanh khdt_store với CSDL tạm trong bộ nhớ: `python test_khdt_store.py` (hoặc pytest)."""
import datetime as dt
import sqlite3

import khdt_core
import khdt_store


def _lop(ten, tu, den, thu_tu):
    nam_t = [f"{y:04d}-{m:02d}" for y, m in khdt_core.cac_thang(dt.date.fromisoformat(tu), dt.date.fromisoformat(den))]
    return dict(mien="NBA", loai_hinh="NVM", loai_hinh_nhan="Ban đầu", stt_goc=str(thu_tu), thu_tu_goc=thu_tu, ten_lop=ten,
                hinh_thuc="Trực tiếp", tu_ngay=tu, den_ngay=den, thoi_gian="", dia_diem="", giao_vien="", so_hv=0,
                doi_tuong="", ghi_chu="", muc="", doi_tac="", dong_goc=thu_tu, thang=nam_t)


def _conn():
    conn = sqlite3.connect(":memory:")
    conn.row_factory = sqlite3.Row
    khdt_store.khoi_tao_bang(conn)
    # KHĐT T11 có lớp học trọn trong T11, lớp vắt T11→T12 và lớp học trọn trong T12
    khdt_store.luu_khdt(conn, 11, 2026, "NBA", "s", "f.xlsx", [
        _lop("Lớp T11", "2026-11-05", "2026-11-05", 1),
        _lop("Lớp vắt tháng", "2026-11-28", "2026-12-02", 2),
        _lop("Lớp T12", "2026-12-10", "2026-12-10", 3)])
    return conn


def test_khdt_thang_lay_du_lop_thang_sau():
    conn = _conn()
    t11 = khdt_store.doc_lop(conn, "NBA", 2026, 11)
    assert [x["ten_lop"] for x in t11] == ["Lớp T11", "Lớp vắt tháng", "Lớp T12"]
    assert not any(x.get("keo_dai") for x in t11)


def test_thang_sau_chua_co_khdt():
    t12 = khdt_store.doc_lop(_conn(), "NBA", 2026, 12)
    assert [x["ten_lop"] for x in t12] == ["Lớp vắt tháng", "Lớp T12"]
    assert all(x["keo_dai"] for x in t12)


def test_gop_nhieu_thang_khong_trung():
    gop = khdt_store.doc_lop_nhieu_thang(_conn(), "NBA", [(2026, 11), (2026, 12)])
    assert [x["ten_lop"] for x in gop] == ["Lớp T11", "Lớp vắt tháng", "Lớp T12"]
    assert not any(x["keo_dai"] for x in gop)


if __name__ == "__main__":
    test_khdt_thang_lay_du_lop_thang_sau()
    test_thang_sau_chua_co_khdt()
    test_gop_nhieu_thang_khong_trung()
    print("OK")
