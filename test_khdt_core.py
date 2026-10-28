"""Kiểm thử nhanh bộ đọc KHĐT (khdt_core) với dòng giả lập theo cấu trúc KHĐT thật: `python test_khdt_core.py` (hoặc pytest)."""
import khdt_core

_TIEU_DE = [["STT", "Khóa đào tạo", "Loại hình đào tạo", "Số lượng học viên", "Đối tượng", "Thời lượng", "", "",
             "Thời gian", "", "Địa điểm", "Giáo viên/ Đối tác", "Ghi chú"],
            ["", "", "", "", "", "EL", "LT", "TH", "Từ ngày", "Đến ngày", "", "", ""]]


def _dong(stt, ten, loai="", sl="", tu="", den="", gv=""):
    return [stt, ten, loai, sl, "", "", "", "", tu, den, "VNBA", gv, ""]


def _doc(*dong):
    return khdt_core.doc_khdt_sheet(_TIEU_DE + [list(d) for d in dong], "NBA", 11, 2026)


def test_dong_la_ma_co_loai_hinh_la_tieu_de_mon():
    """'VI Điều khiển xe…' ghi loại hình nhưng không ngày, con '1 Lý thuyết' + '2 Thực hành' → 1 lớp mang tên môn."""
    lops, canh_bao = _doc(
        _dong("C", "KHU VỰC NỘI BÀI"), _dong("C.3", "ĐÀO TẠO ĐỊNH KỲ"), _dong("C.3.2", "CHUYÊN MÔN NGHIỆP VỤ"),
        _dong("VI", "Điều khiển xe nâng hàng", "Định kỳ"),
        _dong("1", "Lý thuyết + kiểm tra", "", "4", "S3/11", "S3/11", "GV A"),
        _dong("2", "Thực hành + kiểm tra"),
        _dong("", "Thực hành + kiểm tra (nhóm 1)", "", "2hv", "C3/11", "C3/11", "GV B"))
    assert canh_bao == []
    assert [(x["ten_lop"], x["loai_hinh"], x["thoi_gian"]) for x in lops] == [("Điều khiển xe nâng hàng", "DINH_KY", "S3/11/2026")]


def test_dong_khong_stt_co_loai_hinh_la_lop_rieng():
    """'Manual check-in…' không STT nhưng có loại hình + ngày riêng, sau lớp đã có ngày → lớp riêng, không gộp vào lớp trên."""
    lops, _ = _doc(
        _dong("C", "KHU VỰC NỘI BÀI"), _dong("C.5", "ĐÀO TẠO BỒI DƯỠNG KIẾN THỨC"), _dong("C.5.2", "CHUYÊN MÔN NGHIỆP VỤ"),
        _dong("I", "Nghiệp vụ kết sổ chuyến bay - Các hãng"),
        _dong("3", "Hệ thống làm thủ tục", "Bồi dưỡng KT", "7", "6,9,10/11", ""),
        _dong("", "Manual check-in", "Bồi dưỡng KT", "7", "S11/11", "S11/11"))
    assert [(x["ten_lop"], x["thoi_gian"]) for x in lops] == [
        ("Hệ thống làm thủ tục/Nghiệp vụ kết sổ chuyến bay - Các hãng", "6-10/11/2026"),
        ("Manual check-in/Nghiệp vụ kết sổ chuyến bay - Các hãng", "S11/11/2026")]


def test_dong_khong_stt_la_dong_con_khi_lop_tren_chua_co_ngay():
    """'6 Kiểm tra tốt nghiệp' chưa có ngày, dòng 'Tốt nghiệp Lý thuyết' (có loại hình) bên dưới vẫn là dòng con của nó."""
    lops, canh_bao = _doc(
        _dong("C", "KHU VỰC NỘI BÀI"), _dong("C.1", "ĐÀO TẠO BAN ĐẦU - NHÂN VIÊN MỚI"),
        _dong("6", "Kiểm tra tốt nghiệp"),
        _dong("", "Tốt nghiệp Lý thuyết", "Ban đầu", "6", "S14/11", "S14/11"))
    assert canh_bao == [] and [x["ten_lop"] for x in lops] == ["Kiểm tra tốt nghiệp"]


def test_goi_y_sheet_bo_sheet_an_xuong_cuoi():
    ten = ["A-Đào tạo tại TSN", "C-Đào tạo tại NBA xác nhận", "B-Đào tạo tại DAD", "C-Đào tạo tại NBA"]
    assert khdt_core.goi_y_sheet(ten, "NBA") == ["C-Đào tạo tại NBA xác nhận", "C-Đào tạo tại NBA"]
    assert khdt_core.goi_y_sheet(ten, "NBA", an=["C-Đào tạo tại NBA xác nhận"]) == ["C-Đào tạo tại NBA", "C-Đào tạo tại NBA xác nhận"]


if __name__ == "__main__":
    test_dong_la_ma_co_loai_hinh_la_tieu_de_mon()
    test_dong_khong_stt_co_loai_hinh_la_lop_rieng()
    test_dong_khong_stt_la_dong_con_khi_lop_tren_chua_co_ngay()
    test_goi_y_sheet_bo_sheet_an_xuong_cuoi()
    print("OK")
