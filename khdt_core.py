"""Đọc Kế hoạch đào tạo (KHĐT) -> danh sách lớp theo miền × loại hình.

Module thuần Python (không phụ thuộc Streamlit) để dễ kiểm thử.
Quy tắc dựng lớp xem CONTEXT.md mục 3.1b và 6.5b.
"""
import re
import unicodedata
from datetime import date, datetime

# ---------------------------------------------------------------- danh mục
MIEN = {
    "TSN": "Tân Sơn Nhất (TSN)",
    "DAD": "Đà Nẵng (DAD)",
    "NBA": "Nội Bài (NBA)",
    "VNA_DT": "VNA & đối tác",
}

MIEN_FILE = {"TSN": "TSN", "DAD": "DAD", "NBA": "NBA", "VNA_DT": "VNA"}  # dùng đặt tên file (file đối tác đặt tên riêng: DOI TAC <tên>)

# Thứ tự cố định dùng để sắp xếp "theo KHĐT"
LOAI_HINH = {
    "NVM": "Ban đầu - Nhân viên mới",
    "BSND": "Ban đầu - Bổ sung năng định",
    "DINH_KY": "Định kỳ",
    "PHUC_HOI": "Phục hồi",
    "BDKT": "Bồi dưỡng kiến thức",
    "KHAC": "Đào tạo khác",
}
LOAI_HINH_FILE = {  # dùng đặt tên file
    "NVM": "BAN DAU NVM",
    "BSND": "BO SUNG NANG DINH",
    "DINH_KY": "DINH KY",
    "PHUC_HOI": "PHUC HOI",
    "BDKT": "BOI DUONG KT",
    "KHAC": "DAO TAO KHAC",
}
LOAI_HINH_FILE["DOI_TAC"] = "DOI TAC"
MIEN_TACH_KHOA_NVM = ("TSN", "DAD", "NBA")  # NVM của 3 chi nhánh: 1 khóa = 1 file
THU_TU_LOAI = {k: i for i, k in enumerate(LOAI_HINH)}
THU_TU_MIEN = {k: i for i, k in enumerate(MIEN)}

_RE_MA_KHOA = re.compile(r"[A-ZĐ]{2,6}\d{2}\s*-\s*[A-ZĐ]+\s*\d+")


def tach_ma_khoa(doi_tuong):
    """Mã khóa nhân viên mới trong cột 'Đối tượng', vd 'VNBA26- PVHK08' -> 'VNBA26-PVHK08'; không có thì ''."""
    m = _RE_MA_KHOA.search(str(doi_tuong or "").upper())
    return re.sub(r"\s+", "", m.group()) if m else ""


def bo_dau(s):
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.replace("đ", "d").replace("Đ", "D")


def ten_doi_tac(khu_vuc, heads):
    """Tên đối tác của lớp: khu vực 'ĐÀO TẠO ĐỐI TÁC' → tiêu đề mục B.n ('CÔNG TY AGS' -> 'AGS'); không phải đối tác → ''."""
    if "doi tac" not in bo_dau(khu_vuc).lower():
        return ""
    ten = " ".join(str(heads[0]).split()) if heads else ""
    m = re.match(r"(?i)cong\s+ty\s+", bo_dau(ten))
    return (ten[m.end():] if m else ten) or "Khác"


def lam_sach(v):
    """Ô Excel -> chuỗi gọn; None/NaN -> ''."""
    if v is None:
        return ""
    if isinstance(v, float) and v != v:
        return ""
    s = str(v).replace("\xa0", " ").replace("\n", " ").replace("\r", " ")
    s = " ".join(s.split())
    return "" if s.lower() == "nan" else s


# ---------------------------------------------------------------- chọn sheet theo miền
def goi_y_sheet(ten_sheets, mien, an=()):
    """Danh sách sheet gợi ý cho một miền (ưu tiên đứng trước).

    an: các sheet đang ẩn trong file (thường là bản nháp, vd 'C-Đào tạo tại NBA xác nhận') → xếp sau sheet hiển thị.
    """
    ket_qua = []
    for ten in ten_sheets:
        k = bo_dau(ten).lower()
        if mien == "TSN" and "tsn" in k:
            ket_qua.append(ten)
        elif mien == "DAD" and "dad" in k:
            ket_qua.append(ten)
        elif mien == "NBA" and "nba" in k and not re.search(r"\bra\b", k):
            ket_qua.append(ten)  # bỏ qua bản "NBA rà"
        elif mien == "VNA_DT" and "vna" in k:
            ket_qua.append(ten)
    return sorted(ket_qua, key=lambda ten: ten in an)  # sort ổn định: giữ thứ tự trong file


def doan_thang_nam(rows):
    """Tìm 'THÁNG 9/2026' ở các dòng đầu của sheet -> (thang, nam) hoặc (None, None)."""
    for row in rows[:8]:
        for v in row:
            m = re.search(r"th[aá]ng\s*(\d{1,2})\s*/\s*(\d{4})", lam_sach(v), re.IGNORECASE)
            if m:
                return int(m.group(1)), int(m.group(2))
    return None, None


# ---------------------------------------------------------------- ngày
_NGAY = re.compile(
    r"(?<![A-Za-zÀ-ỹ])(?:([SCTsct])\s*)?"
    r"(\d{1,2}(?:\s*[,\-–]\s*\d{1,2})*)\s*/\s*(\d{1,2})(?:\s*/\s*(\d{4}|\d{2}))?(?!\d)"
)
_NGAY_MO_COI = re.compile(r"(?<![\d/])(\d{1,2}(?:\s*[,\-–]\s*\d{1,2})+)(?![\d/])")


def _tao_ngay(day, month, nam_chuoi, thang_ref, nam_ref):
    if nam_chuoi:
        y = int(nam_chuoi)
        if y < 100:
            y += 2000
    else:
        y = nam_ref
        if thang_ref and month - thang_ref >= 7:
            y -= 1
        elif thang_ref and thang_ref - month >= 7:
            y += 1
    try:
        return date(y, month, day)
    except ValueError:
        return None


def _bung_ngay(chuoi):
    """'05,12,19' / '7-9' -> các ngày (lấy mọi mốc, đủ để tìm min/max)."""
    ngay = []
    for phan in re.split(r"\s*,\s*", chuoi):
        for so in re.split(r"\s*[\-–]\s*", phan):
            if so.strip().isdigit():
                ngay.append(int(so))
    return ngay


def phan_tich_ngay(gia_tri, thang_ref=None, nam_ref=None):
    """Đọc một ô thời gian của KHĐT.

    Hỗ trợ: '03/9', 'S3/9', 'C09,10/9', '05,12,19/9', '11-13/9', '14//9', 'Từ 12/9',
    'Trực tiếp: 7-9; 21-22/9', 'Tháng 9', cả kiểu datetime của Excel.
    Trả về (list[date], ca_đầu) — list rỗng nếu không đọc được.
    """
    nam_ref = nam_ref or datetime.now().year
    if isinstance(gia_tri, (datetime, date)):
        d = gia_tri.date() if isinstance(gia_tri, datetime) else gia_tri
        return [d], ""
    text = lam_sach(gia_tri).replace("//", "/")
    if not text:
        return [], ""

    ngay, ca_dau = [], ""
    khoang_da_dung = []
    for m in _NGAY.finditer(text):
        ca, ds_ngay, thang, nam_c = m.groups()
        khoang_da_dung.append(m.span())
        if not ca_dau and ca:
            ca_dau = ca.upper()
        mon = int(thang)
        for day in _bung_ngay(ds_ngay):
            d = _tao_ngay(day, mon, nam_c, thang_ref, nam_ref)
            if d:
                ngay.append(d)

    # đoạn chỉ có ngày, thiếu tháng ('7-9; 21-22/9'): mượn tháng của đoạn có tháng gần nhất
    if ngay:
        con_lai = text
        for a, b in reversed(khoang_da_dung):
            con_lai = con_lai[:a] + " " * (b - a) + con_lai[b:]
        moc = [(m.start(), int(t)) for m in _NGAY.finditer(text) for t in [m.group(3)]]
        for m in _NGAY_MO_COI.finditer(con_lai):
            sau = next((t for pos, t in moc if pos > m.start()), None)
            truoc = next((t for pos, t in reversed(moc) if pos < m.start()), None)
            mon = sau or truoc
            if not mon:
                continue
            for day in _bung_ngay(m.group(1)):
                d = _tao_ngay(day, mon, None, thang_ref, nam_ref)
                if d:
                    ngay.append(d)
        return ngay, ca_dau

    # 'Tháng 9', 'Trong tháng 9' -> cả tháng
    m = re.search(r"th[aá]ng\s*(\d{1,2})", text, re.IGNORECASE)
    if m:
        mon = int(m.group(1))
        d1 = _tao_ngay(1, mon, None, thang_ref, nam_ref)
        if d1:
            nxt = date(d1.year + (mon == 12), mon % 12 + 1, 1)
            return [d1, date.fromordinal(nxt.toordinal() - 1)], ""
    return [], ""


def khoang_ngay(*cac_o, nam_ref=None):
    """(đầu, cuối) của một hoặc nhiều ô/chuỗi thời gian bất kỳ, hoặc None. Dùng cho Tab 3."""
    ngay = []
    for o in cac_o:
        ngay.extend(phan_tich_ngay(o, None, nam_ref)[0])
    return (min(ngay), max(ngay)) if ngay else None


def dinh_dang_thoi_gian(dau, cuoi, ca=""):
    """Hiển thị ở ô 'Thời gian' của file khung: S3/9/2026, 5-10/8/2026, 30/7-6/8/2026."""
    if dau == cuoi:
        return f"{ca}{dau.day}/{dau.month}/{dau.year}"
    if dau.year != cuoi.year:
        return f"{dau.day}/{dau.month}/{dau.year}-{cuoi.day}/{cuoi.month}/{cuoi.year}"
    if dau.month == cuoi.month:
        return f"{dau.day}-{cuoi.day}/{dau.month}/{dau.year}"
    return f"{dau.day}/{dau.month}-{cuoi.day}/{cuoi.month}/{cuoi.year}"


def cac_thang(dau, cuoi):
    """Mọi (năm, tháng) mà khoảng [dau, cuoi] chạm tới."""
    ket_qua, y, m = [], dau.year, dau.month
    while (y, m) <= (cuoi.year, cuoi.month):
        ket_qua.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return ket_qua


# ---------------------------------------------------------------- loại hình
def xac_dinh_loai(heads, loai_cot):
    """Loại hình chuẩn: ưu tiên tiêu đề mục (sâu -> nông), sau đó ô 'Loại hình' của lớp."""
    for t in reversed(heads):
        k = bo_dau(t).lower()
        if "bo sung nang dinh" in k:
            return "BSND"
        if "nhan vien moi" in k:
            return "NVM"
        if "dinh ky" in k:
            return "DINH_KY"
        if "phuc hoi" in k:
            return "PHUC_HOI"
        if "boi duong" in k:
            return "BDKT"
        if "dao tao khac" in k or "nhac lai" in k:
            return "KHAC"
        if "ban dau" in k:
            return "NVM"
    k = bo_dau(loai_cot).lower()
    if "ban dau" in k:
        return "NVM"
    if "dinh ky" in k:
        return "DINH_KY"
    if "phuc hoi" in k:
        return "PHUC_HOI"
    if "boi duong" in k or "bdkt" in k:
        return "BDKT"
    return "KHAC"


def nhan_loai_hien_thi(loai_cot, ma_loai):
    """Nhãn ghi ở 'Loại hình/hình thức đào tạo' — theo KHĐT (ô loại hình), thiếu thì theo danh mục."""
    nhan = lam_sach(loai_cot)
    if not nhan:
        return LOAI_HINH[ma_loai]
    if bo_dau(nhan).lower() in ("bdkt", "boi duong kt"):
        return "Bồi dưỡng kiến thức"
    return nhan


# ---------------------------------------------------------------- giáo viên
_VAI_TRO = r"(?:DG|GT|GPC|TG|TTG)"  # dự giảng / giám thị / giảng phê chuẩn / trợ giảng
_RE_VAI_TRO = re.compile(rf"(?<!\w){_VAI_TRO}(?!\w)")
_RE_NHAN_CHAM = re.compile(
    rf"(?<!\w)(?:{_VAI_TRO}(?:\s*[/+\-]\s*{_VAI_TRO})*|dự giảng|giảng phê chuẩn|trợ giảng)\s*:", re.IGNORECASE)
_RE_CUM_VAI_TRO = re.compile(r"trợ giảng|dự giảng|giảng phê chuẩn", re.IGNORECASE)
_RE_DON_VI_CUOI = re.compile(r"(?:\s+(?:ASOC|TTĐT|CNVN|VNBA|VTSN|VIAGS|VINACS|AGS|TSN|DAD|NBA|SGN|CXR))+$")


def _ten_gv_chinh(dong):
    """Tên GV chính trong một dòng của ô Giáo viên; [] nếu là GV phụ (DG/GT/GPC/TG/TTG/trợ giảng…) hoặc rỗng."""
    if _RE_VAI_TRO.search(dong) or _RE_CUM_VAI_TRO.search(dong):
        return []
    dong = re.sub(r"\([^)]*\)?", " ", dong)             # ghi chú trong ngoặc: ngày, ĐGGV…
    dong = re.sub(r"(?<!\w)ĐGGV(?!\w)", " ", dong)
    dong = re.split(r"\s*[-–]\s*", dong, maxsplit=1)[0]   # '- TTĐT NBA', '-Gv AGS', '- VINACS'
    ten = []
    for phan in dong.split("/"):
        phan = _RE_DON_VI_CUOI.sub("", " ".join(phan.split()))
        if phan:
            ten.append(phan)
    return ten


def tach_giao_vien(o):
    """Danh sách GV chính từ ô 'Giáo viên' GỐC (mỗi người một dòng trong ô, chưa qua lam_sach).

    Bỏ GV phụ ghi chú DG / GT / GPC / TG / TTG / trợ giảng (cuối dòng, trong ngoặc, hoặc sau nhãn 'GPC:', 'DG/GT/GPC :',
    'Dự giảng:'… — từ nhãn trở đi bỏ hết); bỏ ghi chú đơn vị / ngày / ĐGGV nhưng giữ tên.
    """
    if o is None or (isinstance(o, float) and o != o):
        return []
    text = unicodedata.normalize("NFC", str(o)).replace("\xa0", " ")
    ket_qua = []
    for dong in re.split(r"[\n\r;,]+", text):
        m = _RE_NHAN_CHAM.search(dong)
        for ten in _ten_gv_chinh(dong[:m.start()] if m else dong):
            if ten not in ket_qua:
                ket_qua.append(ten)
        if m:
            break
    return ket_qua


# ---------------------------------------------------------------- đọc sheet
_RE_LA_MA = re.compile(r"[IVX]+")
_RE_LA_MA_CHAM = re.compile(r"[IVX]+\.\d+")
_RE_MUC1 = re.compile(r"[A-Z]\.\d+")
_RE_MUC2 = re.compile(r"[A-Z]\.\d+\.\d+")
_RE_KHU_VUC = re.compile(r"[A-Z]")
_RE_LOP = re.compile(r"\d+(\.\d+){0,2}")


def _chuan_stt(v):
    s = lam_sach(v)
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s.rstrip(".")


def _so(v):
    m = re.search(r"\d+", lam_sach(v))
    return int(m.group()) if m else 0


def _nhan_cot(hang_tieu_de, hang_phu):
    """Tìm chỉ số cột theo tiêu đề; thiếu thì dùng vị trí mặc định của mẫu F12."""
    cot = dict(stt=0, ten=1, loai=2, sl=3, dt=4, el=5, lt=6, th=7, tu=8, den=9, dd=10, gv=11, gc=12)
    for i, v in enumerate(hang_tieu_de):
        k = bo_dau(lam_sach(v)).lower()
        if k == "stt":
            cot["stt"] = i
        elif "khoa dao tao" in k:
            cot["ten"] = i
        elif "loai hinh" in k:
            cot["loai"] = i
        elif "so luong" in k:
            cot["sl"] = i
        elif "doi tuong" in k:
            cot["dt"] = i
        elif "dia diem" in k:
            cot["dd"] = i
        elif "giao vien" in k:
            cot["gv"] = i
        elif "ghi chu" in k:
            cot["gc"] = i
    for i, v in enumerate(hang_phu):
        k = bo_dau(lam_sach(v)).lower()
        if k == "el":
            cot["el"] = i
        elif k in ("lt", "tt"):
            cot["lt"] = i
        elif k == "th":
            cot["th"] = i
        elif k.startswith("tu ngay"):
            cot["tu"] = i
        elif k.startswith("den") and "ngay" in k:
            cot["den"] = i
    return cot


def _o(row, i):
    return row[i] if 0 <= i < len(row) else None


def _ten_con_truc_tiep(recs, idx, stt):
    """Tên các dòng con trực tiếp ('3' -> '3.1', '3.2') ngay sau dòng idx; dừng khi sang mục khác."""
    ten = []
    for rec in recs[idx + 1:]:
        if re.fullmatch(re.escape(stt) + r"\.\d+", rec["stt"]):
            ten.append(rec["ten"])
        elif rec["stt"].startswith(stt + ".") or rec["stt"] in ("-", ""):
            continue  # cháu / dòng Nhóm xen giữa
        else:
            break
    return ten


def _la_ly_thuyet_thuc_hanh(ten):
    """Tên bắt đầu bằng 'Lý thuyết…' / 'Thực hành…' (vd 'Lý thuyết + kiểm tra'); 'Kiểm tra lý thuyết cuối khóa' thì không."""
    return bool(re.match(r"\s*(ly thuyet|thuc hanh)", bo_dau(ten).lower()))


# Chỉ các mục BSNĐ có tên bắt đầu như sau mới gộp Lý thuyết + Thực hành thành 1 lớp, tránh gộp nhầm
# (định kỳ, phục hồi, BDKT: gộp mọi môn thuần Lý thuyết + Thực hành — xem _duoc_gop)
TIEN_TO_MUC_GOP = ("dieu khien xe", "van hanh thiet bi", "nghiep vu ve sinh")


def _ten_sach(ten):
    """Tên không dấu, chữ thường, bỏ tiền tố '- ' (vd '- Nhóm 1' -> 'nhom 1')."""
    return re.sub(r"^[\s\-–]+", "", bo_dau(ten).lower())


def _la_dong_lt_th_kt(ten):
    """Dòng của môn thuần Lý thuyết / Thực hành: 'Lý thuyết…', 'Thực hành…', 'Kiểm tra…', 'Nhóm n'."""
    return _ten_sach(ten).startswith(("ly thuyet", "thuc hanh", "kiem tra", "nhom"))


def _duoc_gop(ma_loai, ten_mon):
    if ma_loai in ("DINH_KY", "PHUC_HOI", "BDKT"):
        return True
    return ma_loai == "BSND" and _ten_sach(ten_mon).startswith(TIEN_TO_MUC_GOP)


def _la_ranh_gioi(stt):
    """STT mở đầu một môn / mục / khu vực mới (kết thúc môn đang gộp)."""
    return any(p.fullmatch(stt) for p in (_RE_LA_MA, _RE_LA_MA_CHAM, _RE_MUC2, _RE_MUC1, _RE_KHU_VUC))


def _mon_thuan_lt_th(recs, idx):
    """Các dòng của môn La Mã ở recs[idx] nếu môn chỉ gồm Lý thuyết / Thực hành / Kiểm tra / Nhóm; không thì None."""
    dong = []
    for rec in recs[idx + 1:]:
        if _la_ranh_gioi(rec["stt"]):
            break
        dong.append(rec)
    if not dong or not all(_la_dong_lt_th_kt(d["ten"]) for d in dong):
        return None
    # chỉ gộp khi có CẢ Lý thuyết lẫn Thực hành (môn chỉ có Thực hành giữ tên 'Thực hành + kiểm tra/<môn>')
    if not (any(_ten_sach(d["ten"]).startswith("ly thuyet") for d in dong)
            and any(_ten_sach(d["ten"]).startswith("thuc hanh") for d in dong)):
        return None
    return dong


def _mon_chi_thuc_hanh(recs, idx):
    """Các dòng con (không STT hoặc '-') của môn La Mã ở recs[idx] nếu tất cả đều là 'Thực hành…'; không thì None."""
    dong = []
    for rec in recs[idx + 1:]:
        if _la_ranh_gioi(rec["stt"]):
            break
        dong.append(rec)
    if not dong or not all(d["stt"] in ("", "-") and _ten_sach(d["ten"]).startswith("thuc hanh") for d in dong):
        return None
    return dong


def doc_khdt_sheet(rows, mien, thang, nam):
    """Đọc một sheet KHĐT (list các dòng) -> (danh sách lớp, cảnh báo).

    Mỗi lớp là dict: mien, loai_hinh (mã), loai_hinh_nhan, stt_goc, thu_tu_goc, ten_lop,
    hinh_thuc, tu_ngay/den_ngay (ISO), thoi_gian (hiển thị), dia_diem, giao_vien, so_hv,
    doi_tuong, ghi_chu, muc, dong_goc, thang (list 'YYYY-MM').
    """
    canh_bao = []
    hdr = next((i for i, r in enumerate(rows)
                if any("khoa dao tao" in bo_dau(lam_sach(v)).lower() for v in r)), None)
    if hdr is None:
        raise ValueError("Không tìm thấy dòng tiêu đề 'Khóa đào tạo' trong sheet.")
    cot = _nhan_cot(rows[hdr], rows[hdr + 1] if hdr + 1 < len(rows) else [])

    # chuẩn hóa dòng dữ liệu
    recs = []
    for i in range(hdr + 2, len(rows)):
        r = rows[i]
        stt = _chuan_stt(_o(r, cot["stt"]))
        ten = lam_sach(_o(r, cot["ten"]))
        if not stt and not ten:
            continue
        if stt.lower().startswith("phòng") or ten.lower().startswith("tổng giám đốc"):
            break  # phần chữ ký cuối sheet
        recs.append(dict(dong=i + 1, stt=stt, ten=ten, r=r))

    lops = []
    heads, roman, parents = [], "", {}
    khu_vuc = ""  # tiêu đề khu vực (A/B/C); 'ĐÀO TẠO ĐỐI TÁC' → lớp của đối tác
    muc2 = ""  # tiêu đề mục cấp 2 (C.x.1 / C.x.2 / C.x.3 'CHƯƠNG TRÌNH ĐÀO TẠO BỔ SUNG')
    cur = None

    def ket_thuc_lop():
        nonlocal cur
        if cur is None:
            return
        lop, cur = cur, None
        if not lop["ngay"]:
            canh_bao.append(f"Dòng {lop['dong_goc']}: lớp '{lop['ten_goc']}' không có ngày — bỏ qua.")
            return
        dau = min(d for d, _ in lop["ngay"])
        cuoi = max(d for d, _ in lop["ngay"])
        ca = next((c for d, c in sorted(lop["ngay"], key=lambda x: x[0]) if c), "") if dau == cuoi else ""
        gv = ", ".join(lop["gv"])
        ma_loai = xac_dinh_loai(lop["heads"], lop["loai_cot"])
        # không ghép tên mục La Mã chỉ nơi học / hợp đồng ('Tại VCA', 'NCS - HĐĐT02', 'AGS PL28') vào tên lớp
        roman = "" if lop["doi_tac"] or bo_dau(lop["roman"]).lower().startswith("tai ") else lop["roman"]
        if "bo sung" in bo_dau(lop["muc2"]).lower():
            roman = ""  # mục 'Chương trình đào tạo bổ sung' (C.1.3, C.2.3…): tên lớp chỉ là tên gốc
        if lop["parent"]:
            ten_lop = f"{lop['ten_goc']}/{lop['parent']}"
        elif lop["dotted"] and roman:
            ten_lop = f"{lop['ten_goc']}/{roman}"
        elif (ma_loai in ("NVM", "BSND") or (ma_loai == "BDKT" and any("chuyen mon nghiep vu" in bo_dau(h).lower() for h in lop["heads"])))                 and roman and "tờ trình" not in roman.lower():
            ten_lop = f"{lop['ten_goc']}/{roman}"
        else:
            ten_lop = lop["ten_goc"]
        lops.append(dict(
            mien=mien, loai_hinh=ma_loai, loai_hinh_nhan=nhan_loai_hien_thi(lop["loai_cot"], ma_loai),
            stt_goc=lop["stt"], thu_tu_goc=len(lops) + 1, ten_lop=ten_lop,
            hinh_thuc="Elearning + Trực tiếp" if lop["el"] else "Trực tiếp",
            tu_ngay=dau.isoformat(), den_ngay=cuoi.isoformat(),
            thoi_gian=dinh_dang_thoi_gian(dau, cuoi, ca),
            dia_diem=" + ".join(lop["dd"]), giao_vien=gv,
            so_hv=lop["sl_chinh"] or lop["sl_nhom"], doi_tuong=lop["dt"],
            ghi_chu="; ".join(lop["gc"]), muc=" > ".join(lop["heads"]),
            doi_tac=lop["doi_tac"], dong_goc=lop["dong_goc"],
            thang=[f"{y:04d}-{m:02d}" for y, m in cac_thang(dau, cuoi)],
        ))

    def moi_lop(stt, ten, dong, parent, dotted):
        return dict(stt=stt, ten_goc=ten, dong_goc=dong, dong_cuoi=dong, parent=parent,
                    dotted=dotted, roman=roman, heads=list(heads), loai_cot="", sl_chinh=0,
                    sl_nhom=0, dt="", el=False, ngay=[], gv=[], dd=[], gc=[],
                    doi_tac=ten_doi_tac(khu_vuc, heads), muc2=muc2)

    def cong_du_lieu(lop, r, la_nhom):
        for cell in (_o(r, cot["tu"]), _o(r, cot["den"])):
            ngay, ca = phan_tich_ngay(cell, thang, nam)
            lop["ngay"].extend((d, ca if d == ngay[0] else "") for d in ngay)
            if cell not in (None, "") and not ngay:
                canh_bao.append(f"Dòng {lop['dong_cuoi']}: không đọc được ngày '{lam_sach(cell)}'.")
        for key, c in (("gv", cot["gv"]), ("dd", cot["dd"]), ("gc", cot["gc"])):
            moi = tach_giao_vien(_o(r, c)) if key == "gv" else [lam_sach(_o(r, c))]
            for t in moi:
                if t and t not in lop[key]:
                    lop[key].append(t)
        el = lam_sach(_o(r, cot["el"]))
        if el and el not in ("0", "0.0"):
            lop["el"] = True
        n = _so(_o(r, cot["sl"]))
        if la_nhom:
            lop["sl_nhom"] += n
        elif not lop["sl_chinh"]:
            lop["sl_chinh"] = n
        if not lop["dt"]:
            lop["dt"] = lam_sach(_o(r, cot["dt"]))
        if not lop["loai_cot"]:
            lop["loai_cot"] = lam_sach(_o(r, cot["loai"]))

    for idx, rec in enumerate(recs):
        stt, ten, r = rec["stt"], rec["ten"], rec["r"]
        if cur is not None and cur.get("gop_het") and not _la_ranh_gioi(stt):
            # môn thuần Lý thuyết + Thực hành: mọi dòng (kể cả Kiểm tra, Nhóm) cộng vào cùng một lớp
            cur["dong_cuoi"] = rec["dong"]
            cong_du_lieu(cur, r, _ten_sach(ten).startswith("nhom"))
            continue
        if _RE_LA_MA_CHAM.fullmatch(stt):  # I.1, II.2
            ket_thuc_lop()
            heads = heads[:3] + [ten]
        elif _RE_LA_MA.fullmatch(stt) and (
                any(lam_sach(_o(r, cot[k])) for k in ("tu", "den"))
                or (lam_sach(_o(r, cot["loai"]))
                    and not (idx + 1 < len(recs) and _RE_LOP.fullmatch(recs[idx + 1]["stt"])))):
            # số La Mã nhưng mang dữ liệu lớp (vd 'III Điều khiển xe… - Thực hành thực tế' + các Nhóm);
            # chỉ ghi loại hình, không ngày mà ngay sau là lớp đánh số (1 Lý thuyết…, 2 Thực hành…) thì là tiêu đề môn
            ket_thuc_lop()
            roman, parents = ten, {}
            heads = heads[:2] + [ten]
            cur = moi_lop(stt, ten, rec["dong"], "", False)
            cur["roman"] = ""  # tên lớp là chính môn này, không ghép tên môn La Mã đứng trước
            cong_du_lieu(cur, r, False)
            con = _mon_chi_thuc_hanh(recs, idx)
            if con and xac_dinh_loai(heads, cur["loai_cot"]) == "BDKT":
                # môn BDKT chỉ có các dòng Thực hành (+ kiểm tra, nhóm) → 'Thực hành + kiểm tra/<môn>'
                cur["parent"] = ten
                cur["ten_goc"] = "Thực hành + kiểm tra" if any("kiem tra" in bo_dau(d["ten"]).lower() for d in con) else "Thực hành"
        elif _RE_LA_MA.fullmatch(stt):  # I, II (môn / nghiệp vụ)
            ket_thuc_lop()
            roman, parents = ten, {}
            heads = heads[:2] + [ten]
            mon = _mon_thuan_lt_th(recs, idx)
            if mon:  # môn có cả Lý thuyết lẫn Thực hành (+ kiểm tra, nhóm) → 1 lớp, tên = tên môn
                loai_cot = next((lam_sach(_o(d["r"], cot["loai"])) for d in mon if lam_sach(_o(d["r"], cot["loai"]))),
                                lam_sach(_o(r, cot["loai"])))  # dòng con trống loại hình → lấy của dòng môn
                if _duoc_gop(xac_dinh_loai(heads, loai_cot), ten):
                    cur = moi_lop(stt, ten, rec["dong"], "", False)
                    cur["gop_het"] = True
                    cur["roman"] = ""  # tên lớp chính là tên môn, không ghép thêm
                    cong_du_lieu(cur, r, False)
        elif _RE_MUC2.fullmatch(stt):  # C.2.1
            ket_thuc_lop()
            roman, parents, muc2 = "", {}, ten
            heads = heads[:1] + [ten]
        elif _RE_MUC1.fullmatch(stt):  # C.2
            ket_thuc_lop()
            roman, parents, heads, muc2 = "", {}, [ten], ""
        elif _RE_KHU_VUC.fullmatch(stt):  # A / B / C (khu vực)
            ket_thuc_lop()
            roman, parents, heads, khu_vuc, muc2 = "", {}, [], ten, ""
        elif _RE_LOP.fullmatch(stt):
            if cur is not None and cur.get("gop") and stt.startswith(cur["stt"] + "."):
                # dòng con Lý thuyết / Thực hành của lớp gộp: cộng dữ liệu vào cùng một lớp
                cur["dong_cuoi"] = rec["dong"]
                cong_du_lieu(cur, r, False)
                continue
            ket_thuc_lop()
            co_ngay = any(lam_sach(_o(r, cot[k])) for k in ("tu", "den"))
            ke_tiep = recs[idx + 1]["stt"] if idx + 1 < len(recs) else ""
            # dòng không có ngày mà ngay sau là các dòng con (3 → 3.1, 3.2) là lớp cha / lớp gộp,
            # kể cả khi ô loại hình ('Ban đầu') của nó có dữ liệu
            if not co_ngay and ke_tiep.startswith(stt + "."):
                con = _ten_con_truc_tiep(recs, idx, stt)
                gop_kt = _duoc_gop(xac_dinh_loai(heads, ""), ten)
                if con and all(_la_ly_thuyet_thuc_hanh(t) or (gop_kt and _ten_sach(t).startswith("kiem tra")) for t in con):
                    # khóa chỉ gồm các lớp nhỏ Lý thuyết / Thực hành → gộp thành 1 lớp (tên = tên khóa)
                    cur = moi_lop(stt, ten, rec["dong"], "", False)
                    cur["gop"] = True
                    cong_du_lieu(cur, r, False)
                    continue
                parents[stt] = ten  # lớp cha: chỉ làm tiêu đề, không tạo sheet
                continue
            dotted = "." in stt
            parent = ""
            if dotted:
                cha = stt.rsplit(".", 1)[0]
                parent = parents.get(cha, "")
            cur = moi_lop(stt, ten, rec["dong"], parent, dotted)
            cong_du_lieu(cur, r, False)
        elif not stt and lam_sach(_o(r, cot["loai"])) and not _la_dong_lt_th_kt(ten) \
                and any(lam_sach(_o(r, cot[k])) for k in ("tu", "den")) and (cur is None or cur["ngay"]):
            # lớp riêng chỉ thiếu STT: có loại hình + ngày của nó, lớp đứng trên đã đủ ngày (vd 'Manual check-in…' sau
            # lớp 3 'Hệ thống làm thủ tục…'); lớp trên chưa có ngày ('6 Kiểm tra tốt nghiệp') thì vẫn là dòng con của nó
            ket_thuc_lop()
            cur = moi_lop("", ten, rec["dong"], "", False)
            cong_du_lieu(cur, r, False)
        elif stt == "-" or (not stt and (ten.lower().startswith("nhóm")
                                          or any(lam_sach(_o(r, cot[k])) for k in ("tu", "den")))):
            # dòng con của lớp hiện tại: '-', 'Nhóm n', hoặc dòng không STT nhưng có ngày (vd 'Lý thuyết + kiểm tra')
            if cur is None:
                canh_bao.append(f"Dòng {rec['dong']}: '{ten}' không thuộc lớp nào.")
                continue
            cur["dong_cuoi"] = rec["dong"]
            cong_du_lieu(cur, r, ten.lower().startswith("nhóm"))
        else:
            if any(lam_sach(_o(r, cot[k])) for k in ("tu", "den")):
                canh_bao.append(f"Dòng {rec['dong']}: dòng '{ten}' có ngày nhưng không nhận dạng được cấu trúc.")
    ket_thuc_lop()
    return lops, canh_bao


# ---------------------------------------------------------------- lọc / sắp xếp
def chia_file_xuat(lops, mien, muc_chon=None, xuat_vna=True, xuat_doi_tac=True):
    """Chia các lớp (đã sắp xếp) của MỘT miền thành các file xuất.

    Trả về list dict(loai, hau_to, lops), thứ tự lớp trong mỗi file giữ nguyên thứ tự đầu vào:
    - TSN/DAD/NBA: mỗi loại hình (trong muc_chon) 1 file; NVM tách 1 khóa (mã khóa ở cột Đối tượng) = 1 file,
      lớp không có mã khóa gom 1 file;
    - VNA & đối tác: phần VNA (lớp không thuộc đối tác) tách theo loại hình như các miền khác nhưng NVM không tách khóa;
      mỗi đối tác 1 file (loai='DOI_TAC', đủ mọi loại hình).
    """
    ket_qua = []
    la_vna_dt = mien == "VNA_DT"
    for ma in LOAI_HINH:
        if (la_vna_dt and not xuat_vna) or (muc_chon is not None and ma not in muc_chon):
            continue
        nhom = [x for x in lops if x["loai_hinh"] == ma and not (la_vna_dt and x.get("doi_tac"))]
        if not nhom:
            continue
        if ma == "NVM" and mien in MIEN_TACH_KHOA_NVM:
            theo_khoa = {}
            for x in nhom:
                theo_khoa.setdefault(tach_ma_khoa(x.get("doi_tuong")), []).append(x)
            for khoa, ds in sorted(theo_khoa.items(), key=lambda kv: (kv[0] == "", kv[0])):
                ket_qua.append(dict(loai=ma, hau_to=khoa or "KHONG MA KHOA", lops=ds))
        else:
            ket_qua.append(dict(loai=ma, hau_to="", lops=nhom))
    if la_vna_dt and xuat_doi_tac:
        theo_dt = {}
        for x in lops:
            if x.get("doi_tac"):
                theo_dt.setdefault(x["doi_tac"], []).append(x)
        for ten, ds in theo_dt.items():
            ket_qua.append(dict(loai="DOI_TAC", hau_to=ten, lops=ds))
    return ket_qua


def sap_xep(lops, che_do):
    """che_do: 'khdt' = thứ tự gốc trong KHĐT; 'thoi_gian' = ngày tăng dần."""
    if che_do == "thoi_gian":
        return sorted(lops, key=lambda x: (x["tu_ngay"], x["den_ngay"],
                                           THU_TU_MIEN.get(x["mien"], 99),
                                           THU_TU_LOAI.get(x["loai_hinh"], 99), x["thu_tu_goc"]))
    # theo KHĐT: đúng thứ tự dòng trong KHĐT (gộp nhiều tháng: hết KHĐT tháng trước mới đến tháng sau);
    # lớp kéo dài từ KHĐT tháng trước xếp sau cùng
    return sorted(lops, key=lambda x: (bool(x.get("keo_dai")), x.get("k_nam") or 0, x.get("k_thang") or 0,
                                       THU_TU_MIEN.get(x["mien"], 99), x["thu_tu_goc"]))


def nhan_thang(cac_thang):
    """list (nam, thang) → 'T10-2026', 'T10+11-2026', khác năm 'T12-2026+T01-2027' (tên file / tiêu đề)."""
    theo_nam = {}
    for nam, thang in sorted(set(cac_thang)):
        theo_nam.setdefault(nam, []).append(thang)
    return "+".join(f"T{'+'.join(f'{t:02d}' for t in ts)}-{nam}" for nam, ts in theo_nam.items())
