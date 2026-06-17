import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
import io
import zipfile
import re
from datetime import datetime, timedelta

# ==========================================
# CẤU HÌNH GIAO DIỆN
# ==========================================
st.set_page_config(page_title="🛠️ Công cụ Xử lý Dữ liệu Đào tạo VIAGS", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    .status-box { padding: 20px; border-radius: 10px; border: 1px solid #dee2e6; background-color: white; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 📑 HÀM TIỆN ÍCH DÙNG CHUNG
# ==========================================
def remove_vietnamese_accents(s):
    s = str(s)
    s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
    s = re.sub(r'[ÈÉẸẺẼÊỀẾỆỂỄ]', 'E', s)
    s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
    s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)
    s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
    s = re.sub(r'[ìíịỉĩ]', 'i', s)
    s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
    s = re.sub(r'[đ]', 'd', s)
    s = re.sub(r'[ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶ]', 'A', s)
    s = re.sub(r'[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]', 'O', s)
    s = re.sub(r'[ÙÚỤỦŨƯỪỨỰỬỮ]', 'U', s)
    s = re.sub(r'[ÌÍỊỈĨ]', 'I', s)
    s = re.sub(r'[YÝỴỶỸ]', 'Y', s)
    s = re.sub(r'[Đ]', 'D', s)
    return s

def chuan_hoa_thoi_gian(tg_str):
    tg_str = str(tg_str).strip()
    tg_str = re.sub(r'^[CSTcst](?=\d)', '', tg_str) 
    tg_str = re.sub(r'\s*-\s*', ' - ', tg_str)
    return tg_str

def parse_date_range(thoi_gian_str):
    tg_clean = re.sub(r'^[CSTcst]', '', str(thoi_gian_str)).strip()
    tg_clean = tg_clean.replace('.', '/').replace('\u2013', '-').replace('\u2014', '-')
    parts = re.split(r'\s*(?:[–—-])\s*', tg_clean)

    full_dates = []
    day_only_parts = []
    for idx, part in enumerate(parts):
        part = part.strip()
        if not part:
            continue
        match_full = re.match(r'^(\d{1,2})\s*[\/\.-]\s*(\d{1,2})(?:\s*[\/\.-]\s*(\d{2,4}))?$', part)
        if match_full:
            day = int(match_full.group(1))
            month = int(match_full.group(2))
            year = int(match_full.group(3)) if match_full.group(3) else datetime.now().year
            if year < 100:
                year += 2000
            try:
                full_dates.append((idx, datetime(year, month, day)))
            except ValueError:
                continue
            continue
        match_day_only = re.match(r'^(\d{1,2})$', part)
        if match_day_only:
            day_only_parts.append((idx, int(match_day_only.group(1))))

    if full_dates:
        dates = [d for _, d in full_dates]
        for idx, day in day_only_parts:
            reference = next((d for j, d in full_dates if j > idx), None)
            if reference is None:
                reference = next((d for j, d in reversed(full_dates) if j < idx), None)
            if reference is None:
                reference = datetime.now()
            try:
                dates.append(datetime(reference.year, reference.month, day))
            except ValueError:
                continue
        dates.sort()
        return dates[0], dates[-1]

    if day_only_parts:
        day = day_only_parts[0][1]
        reference = datetime.now()
        try:
            date_val = datetime(reference.year, reference.month, day)
            return date_val, date_val
        except ValueError:
            return None, None

    return None, None


def week_labels_from_range(start_date, end_date):
    if not start_date or not end_date:
        return []
    if start_date > end_date:
        start_date, end_date = end_date, start_date

    start_week = start_date - timedelta(days=start_date.weekday())
    end_week = end_date - timedelta(days=end_date.weekday())
    labels = []
    current_week = start_week
    while current_week <= end_week:
        week_num = current_week.isocalendar()[1]
        week_end = current_week + timedelta(days=6)
        if current_week.month == week_end.month:
            labels.append(f"Tuần {week_num} ({current_week.day:02d}-{week_end.day:02d}/{current_week.month})")
        else:
            labels.append(f"Tuần {week_num} ({current_week.day:02d}/{current_week.month}-{week_end.day:02d}/{week_end.month})")
        current_week += timedelta(days=7)
    return labels


def clean_header(val):
    return "".join(str(val).lower().split())

def read_excel_values_only(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith('.csv'):
        try: return {"CSV": pd.read_csv(uploaded_file, header=None, encoding='utf-8')}
        except:
            uploaded_file.seek(0)
            return {"CSV": pd.read_csv(uploaded_file, header=None, encoding='utf-8-sig')}
    elif name.endswith('.xls'):
        return pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    else:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        res = {}
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            data = [list(row) for row in ws.iter_rows(values_only=True)]
            res[sheet_name] = pd.DataFrame(data)
        return res

# ==========================================
# CORE LOGIC BÓC TÁCH KHĐT & TRỘN HỌC VIÊN
# ==========================================
def doc_khdt_gom_theo_tuan(file_khdt):
    """BƯỚC 1: Đọc KHĐT và xuất ra list các lớp theo tuần (Chỉ tạo khung)"""
    file_khdt.seek(0)
    wb_khdt = pd.ExcelFile(file_khdt)
    
    sheet_khdt_name = next((n for n in wb_khdt.sheet_names if "vnba" in n.lower() or "nba" in n.lower()), None)
    if not sheet_khdt_name:
        valid_sheets = [n for n in wb_khdt.sheet_names if "đào tạo" in n.lower()]
        sheet_khdt_name = valid_sheets[0] if valid_sheets else wb_khdt.sheet_names[0]
        
    df_khdt_raw = pd.read_excel(wb_khdt, sheet_name=sheet_khdt_name, header=None)
    header_row_idx = next((idx for idx, row in df_khdt_raw.iterrows() if any("khóa đào tạo" in str(s).lower() or "khoa dao tao" in str(s).lower() for s in row.values)), 4)
    df_khdt_data = pd.read_excel(wb_khdt, sheet_name=sheet_khdt_name, skiprows=header_row_idx)
    
    def is_roman(s): return bool(re.match(r'^(I{1,3}|IV|V|VI{1,3}|IX|X)$', str(s).strip(), re.IGNORECASE))
        
    classes, current_roman_course, current_category, current_class = [], "", "", None
    dict_theo_tuan = {}
    
    for _, row in df_khdt_data.iterrows():
        if len(row) < 12: continue
        stt, ten = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
        loai_hinh, el_val = str(row.iloc[2]).strip(), str(row.iloc[5]).strip()
        tu_ngay, den_ngay = str(row.iloc[8]).strip(), str(row.iloc[9]).strip()
        dia_diem, gv = str(row.iloc[10]).strip(), str(row.iloc[11]).strip()
        
        if pd.isna(row.iloc[1]) or ten == "" or ten == "nan" or "tổng cộng" in ten.lower(): continue
            
        ten_lower = ten.lower()
        if not stt.isdigit() and not is_roman(stt):
            if "đào tạo ban đầu" in ten_lower: current_category, current_roman_course = "Ban đầu", ""
            elif "đào tạo định kỳ" in ten_lower or "đào tạo nhắc lại" in ten_lower: current_category, current_roman_course = "Định kỳ", ""
            elif "đào tạo phục hồi" in ten_lower: current_category, current_roman_course = "Phục hồi", ""
            elif "bồi dưỡng" in ten_lower or "đào tạo bổ sung" in ten_lower or "đào tạo khác" in ten_lower: current_category, current_roman_course = "Bồi dưỡng kiến thức", ""
        
        if is_roman(stt) and stt != "" and stt != "nan":
            current_roman_course = ten; continue
            
        if stt.isdigit():
            if current_class: classes.append(current_class)
            final_loai_hinh = loai_hinh if loai_hinh != "nan" and loai_hinh != "" else current_category
            if final_loai_hinh.lower() == "bồi dưỡng kt" or final_loai_hinh.lower() == "bdkt":
                final_loai_hinh = "Bồi dưỡng kiến thức"
            has_elearning = el_val and el_val != "nan" and el_val.strip() != ""
            current_class = {
                "ten_goc": ten, "khoa_hoc_la_ma": current_roman_course, "loai_hinh": final_loai_hinh,
                "hinh_thuc": "Elearning + Trực tiếp" if has_elearning else "Trực tiếp",
                "dia_diem": dia_diem if dia_diem != "nan" else "", "dia_diem_lt": "", "dia_diem_th": "",
                "gv_lt": gv if gv != "nan" else "", "tu_ngay_goc": tu_ngay if tu_ngay != 'nan' else "",
                "den_ngay_goc": den_ngay if den_ngay != 'nan' else "", "lt_dates": [], "th_dates": [], "nhom_dates": []
            }
            if current_class["tu_ngay_goc"]: current_class["lt_dates"].append(current_class["tu_ngay_goc"])
            if current_class["den_ngay_goc"]: current_class["th_dates"].append(current_class["den_ngay_goc"])
                
        elif current_class is not None:
            row_text_lower = (stt + " " + ten).lower()
            
            # --- CODE MỚI THÊM VÀO ĐÂY: Quét bù cột EL ở dòng phụ ---
            if el_val and el_val != "nan" and el_val.strip() != "":
                current_class["hinh_thuc"] = "Elearning + Trực tiếp"
            # --------------------------------------------------------
            
            if "nhóm" in row_text_lower:
                if tu_ngay != 'nan' and den_ngay != 'nan': current_class["nhom_dates"].append((tu_ngay, den_ngay))
            elif "lý thuyết" in row_text_lower or "lt" in row_text_lower.split() or "lý thuyết" in str(row.iloc[2]).lower():
                if tu_ngay != 'nan': current_class["lt_dates"].append(tu_ngay)
                if gv != 'nan' and not current_class["gv_lt"]: current_class["gv_lt"] = gv 
                if dia_diem != 'nan' and dia_diem: current_class["dia_diem_lt"] = dia_diem
            elif "thực hành" in row_text_lower or "th" in row_text_lower.split() or "thực hành" in str(row.iloc[2]).lower():
                if den_ngay != 'nan': current_class["th_dates"].append(den_ngay)
                elif tu_ngay != 'nan': current_class["th_dates"].append(tu_ngay)
                if dia_diem != 'nan' and dia_diem: current_class["dia_diem_th"] = dia_diem

    if current_class: classes.append(current_class)

    for c in classes:
        loai = str(c["loai_hinh"]).lower()
        ten_final = c["ten_goc"]
        thoi_gian = ""
        
        if "ban đầu" in loai or "ban dau" in loai:
            if c["khoa_hoc_la_ma"] and "tờ trình" not in c["khoa_hoc_la_ma"].lower():
                ten_final = f"{c['ten_goc']}/{c['khoa_hoc_la_ma']}"
            if c["nhom_dates"]:
                start, end = c["nhom_dates"][0][0], c["nhom_dates"][-1][1]
                thoi_gian = f"{start} - {end}" if start != end else start
            else:
                s, e = c["tu_ngay_goc"], c["den_ngay_goc"]
                thoi_gian = f"{s} - {e}" if s and e and s != e else (s or e)
            if not c["loai_hinh"]: c["loai_hinh"] = "Ban đầu"
        else:
            s = c["lt_dates"][0] if c["lt_dates"] else c["tu_ngay_goc"]
            e = c["th_dates"][-1] if c["th_dates"] else c["den_ngay_goc"]
            s_clean, e_clean = re.sub(r'^[CSTcst]', '', str(s)).strip(), re.sub(r'^[CSTcst]', '', str(e)).strip()
            if s_clean and e_clean and s_clean == e_clean: thoi_gian = s_clean
            elif s_clean and e_clean: thoi_gian = f"{s} - {e}" 
            else: thoi_gian = s_clean or e_clean
        
        if not thoi_gian: continue 
        
        dd_lt, dd_th = c["dia_diem_lt"] or c["dia_diem"], c["dia_diem_th"] or c["dia_diem"]
        final_dia_diem = f"{dd_lt} + {dd_th}" if dd_lt and dd_th and dd_lt != dd_th else (dd_lt or dd_th or c["dia_diem"])
        
        t_chuan = chuan_hoa_thoi_gian(thoi_gian)
        start_date, end_date = parse_date_range(t_chuan)
        week_labels = week_labels_from_range(start_date, end_date) if start_date and end_date else ["Tuần_Khác"]

        lop_info = {
            "ten_lop": ten_final, "thoi_gian": thoi_gian, "hinh_thuc": c["hinh_thuc"],
            "loai_hinh": c["loai_hinh"], "dia_diem": final_dia_diem, "giao_vien": c["gv_lt"]
        }

        for label in week_labels:
            if label not in dict_theo_tuan:
                dict_theo_tuan[label] = []
            dict_theo_tuan[label].append(lop_info)
        
    return dict_theo_tuan

def tao_file_excel_mot_tuan(tuan_name, dslop):
    """Tạo Form Excel Template rỗng"""
    wb = openpyxl.Workbook()
    ws_index = wb.active
    ws_index.title = "MucLuc"
    ws_index.column_dimensions['A'].width = 8
    ws_index.column_dimensions['B'].width = 50
    ws_index.column_dimensions['C'].width = 25
    ws_index.column_dimensions['D'].width = 20
    ws_index.column_dimensions['E'].width = 25
    ws_index.column_dimensions['F'].width = 20
    
    headers = ["STT", "TÊN LỚP (BẤM VÀO ĐỂ TỚI SHEET)", "LOẠI HÌNH/HÌNH THỨC", "THỜI GIAN", "ĐỊA ĐIỂM", "GIÁO VIÊN"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws_index.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
    
    for idx, lop in enumerate(dslop):
        row_idx = idx + 2
        ws_index[f'A{row_idx}'] = idx + 1
        
        safe_name = re.sub(r'[\\/*?:\[\]]', '', lop["ten_lop"])[:30]
        if not safe_name: safe_name = f"Lop_Hoc"
        while safe_name in wb.sheetnames: safe_name = safe_name[:27] + f"_{idx}"
            
        ws = wb.create_sheet(title=safe_name)
        ws['A1'] = "⬅️ Trở về Mục lục"
        ws['A1'].hyperlink = "#'MucLuc'!A1"
        ws['A1'].font = Font(color="0563C1", underline="single", bold=True)
        
        link_cell = ws_index[f'B{row_idx}']
        link_cell.value = lop["ten_lop"]
        link_cell.hyperlink = f"#'{safe_name}'!A1"
        link_cell.font = Font(color="0563C1", underline="single")
        
        ws_index[f'C{row_idx}'] = f"{lop['loai_hinh']}/{lop['hinh_thuc']}"
        ws_index[f'D{row_idx}'] = lop["thoi_gian"]
        ws_index[f'E{row_idx}'] = lop["dia_diem"]
        ws_index[f'F{row_idx}'] = lop["giao_vien"]
        
        ws['D7'], ws['D9'], ws['D10'], ws['D11'] = lop["ten_lop"], lop["thoi_gian"], lop["dia_diem"], lop["giao_vien"]
        ws['B7'] = f"- Môn học/Khóa học:"
        ws['B8'] = f"- Loại hình/hình thức đào tạo: {lop['loai_hinh']}/{lop['hinh_thuc']}"
        ws['B9'] = f"- Thời gian:"
        ws['B10'] = f"- Địa điểm:"
        ws['B11'] = f"- Giáo viên:"
        ws['L11'] = "" 
        
        headers = ["STT", "Mã NV", "Họ tên", "Đơn vị", "Ghi chú"]
        for c_idx, h in enumerate(headers, start=1):
            ws.cell(row=13, column=c_idx, value=h).font = Font(bold=True)
            
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 30
        ws.column_dimensions['D'].width = 20
        
        for r_idx in range(14, 29):
            ws.cell(row=r_idx, column=1, value=r_idx-13)
            
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

def tao_file_excel_tu_dict(dict_theo_tuan):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for tuan, dslop in dict_theo_tuan.items():
            excel_buffer = tao_file_excel_mot_tuan(tuan, dslop)
            safe_tuan = str(tuan).replace("/", "-").replace(":", "")
            zf.writestr(f"Template_DSLop_{safe_tuan}.xlsx", excel_buffer.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

def doc_dshv_ra_list(file_dshv):
    """BƯỚC 2A: Đọc file DS T6 (File HV) bóc tách list Học viên"""
    file_dshv.seek(0)
    wb_dshv = pd.ExcelFile(file_dshv)
    ds_lop_hv = []
    
    for sheetname in wb_dshv.sheet_names:
        if remove_vietnamese_accents(sheetname).lower() in ["mucluc", "sheet1"]: continue
        df_sheet = pd.read_excel(wb_dshv, sheet_name=sheetname, header=None)
        header_row_hv = next((idx for idx, row in df_sheet.iterrows() if any("khóa học" in str(s).lower() or "khoa hoc" in str(s).lower() for s in row.values)), None)
        
        if header_row_hv is None: continue
        current_hv_class = None
        
        for idx in range(header_row_hv + 1, len(df_sheet)):
            row = df_sheet.iloc[idx]
            val_khoa_hoc = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            
            if val_khoa_hoc != "" and val_khoa_hoc != "nan" and not val_khoa_hoc.startswith("C."):
                tu_ngay_hv = str(row.iloc[6]).strip() if len(row) > 6 else ""
                den_ngay_hv = str(row.iloc[7]).strip() if len(row) > 7 else ""
                tg_hv = f"{tu_ngay_hv} - {den_ngay_hv}" if tu_ngay_hv != den_ngay_hv and den_ngay_hv and den_ngay_hv != "nan" else tu_ngay_hv
                tg_hv_chuan = chuan_hoa_thoi_gian(tg_hv)
                ten_lop_hv = val_khoa_hoc.split("\n")[0].strip()
                
                current_hv_class = {"ten_lop": ten_lop_hv, "thoi_gian": tg_hv_chuan, "hoc_vien": []}
                ds_lop_hv.append(current_hv_class)
                
            elif current_hv_class is not None:
                ma_nv = str(row.iloc[2]).strip() if len(row) > 2 else ""
                ho_ten = str(row.iloc[3]).strip() if len(row) > 3 else ""
                don_vi = str(row.iloc[4]).strip() if len(row) > 4 else ""
                if ho_ten and ho_ten != "nan" and ho_ten != "HỌ VÀ TÊN":
                    current_hv_class["hoc_vien"].append({"manv": ma_nv if ma_nv != "nan" else "", "hoten": ho_ten, "donvi": don_vi if don_vi != "nan" else ""})
                    
    return ds_lop_hv

def nhoi_hoc_vien_vao_template(file_template, ds_lop_hv):
    file_template.seek(0)
    wb = openpyxl.load_workbook(file_template)
    
    for sheetname in wb.sheetnames:
        if remove_vietnamese_accents(sheetname).lower() == "mucluc": continue
        ws = wb[sheetname]
        
        ten_lop_kh = str(ws['D7'].value or "")
        tg_kh = str(ws['D9'].value or "")

        # Map sheet name to standardized 'loại hình'
        s_name_clean = remove_vietnamese_accents(sheetname).lower()
        def map_sheet_loai(s):
            if any(k in s for k in ['ban dau', 'bandau', 'bsn', 'bsnd']):
                return 'Ban đầu'
            if 'dinh ky' in s or 'dinhky' in s:
                return 'Định kỳ'
            if 'phuc hoi' in s or 'phuchoi' in s:
                return 'Phục hồi'
            # BDKT or any other sheet -> Bồi dưỡng KT
            return 'Bồi dưỡng KT'

        mapped_loai = map_sheet_loai(s_name_clean)
        # preserve existing hình thức (if present in B8) when updating B8
        existing_b8 = str(ws.get('B8').value or '')
        hinh_thuc = ''
        m = re.search(r':\s*(.*)', existing_b8)
        if m:
            after = m.group(1)
            parts = [p.strip() for p in after.split('/')]
            if len(parts) > 1:
                hinh_thuc = parts[1]
            else:
                hinh_thuc = parts[0]

        if hinh_thuc:
            ws['B8'] = f"- Loại hình/hình thức đào tạo: {mapped_loai}/{hinh_thuc}"
        else:
            ws['B8'] = f"- Loại hình/hình thức đào tạo: {mapped_loai}"

        key_kh_ten = remove_vietnamese_accents(ten_lop_kh).lower()
        key_kh_tg = remove_vietnamese_accents(tg_kh).lower()

        matched_hv = []
        for hv_class in ds_lop_hv:
            key_hv_ten = remove_vietnamese_accents(hv_class["ten_lop"]).lower()
            key_hv_tg = remove_vietnamese_accents(hv_class["thoi_gian"]).lower()
            
            if (key_kh_ten in key_hv_ten or key_hv_ten in key_kh_ten) and (key_kh_tg in key_hv_tg or key_hv_tg in key_kh_tg):
                matched_hv = hv_class["hoc_vien"]
                break
                
        if matched_hv:
            # Dọn dẹp hàng cũ từ dòng 14 đến 40 để tránh đè dữ liệu rác
            for r in range(14, max(14 + len(matched_hv), 30)):
                for c in range(1, 6): ws.cell(row=r, column=c).value = ""
            # Điền danh sách mới
            for i, hv in enumerate(matched_hv):
                r_idx = 14 + i
                ws.cell(row=r_idx, column=1, value=i+1)
                ws.cell(row=r_idx, column=2, value=hv["manv"])
                ws.cell(row=r_idx, column=3, value=hv["hoten"])
                ws.cell(row=r_idx, column=4, value=hv["donvi"])

    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# ==========================================
# GIAO DIỆN CHÍNH (STREAMLIT TABS)
# ==========================================
st.title("🚀 Hệ thống Trộn & Đối chiếu Dữ liệu Đào tạo VIAGS")

# Chia ứng dụng thành 3 tính năng riêng biệt sạch sẽ
tab_doi_chieu, tab_tao_khung, tab_nhoi_hv = st.tabs([
    "🔍 1. ĐỐI CHIẾU DANH SÁCH SƠ TUYỂN", 
    "📄 2. TẠO KHUNG TỪ KHĐT", 
    "🧑‍🎓 3. TỰ ĐỘNG NHỒI HỌC VIÊN"
])

# --- TAB 1: ĐỐI CHIẾU DANH SÁCH (GIỮ NGUYÊN HOÀN TOÀN TÍNH NĂNG GỐC CỦA APP SO KHỚP) ---
with tab_doi_chieu:
    st.info("💡 Hướng dẫn: Radar sẽ tự động quét tiêu đề (Mã NV, Họ tên) dù ở bất kỳ vị trí nào trong file.")
    
    col_file1, col_file2 = st.columns(2)
    with col_file1:
        file_dsnv = st.file_uploader("1. File Danh sách nhân viên gốc (DSNV)", type=["xlsx", "xls", "csv"])
    with col_file2:
        file_dshv = st.file_uploader("2. File Danh sách học viên cần kiểm tra", type=["xlsx", "xls", "csv"])

    if st.button("🚀 Bắt đầu quét đối chiếu", type="primary"):
        if not file_dsnv or not file_dshv:
            st.warning("⚠️ Vui lòng tải lên đầy đủ cả 2 file!")
        else:
            with st.spinner("Đang quét dữ liệu..."):
                try:
                    # 1. Xử lý File Gốc
                    df_dsnv_all = read_excel_values_only(file_dsnv)
                    dict_nv_chuan = {}
                    kw_ma = ['mãnv', 'mnv', 'manv', 'mãsốnv', 'mãnhânviên', 'staffid']
                    kw_ten = ['họvàtên', 'họtên', 'fullname']

                    for s_name, df_s in df_dsnv_all.items():
                        c_ma, c_ten = -1, -1
                        for idx, row in df_s.iterrows():
                            row_raw = [str(x).strip() if x is not None else "" for x in row.values]
                            row_cleaned = [clean_header(x) for x in row_raw]
                            
                            tmp_ma = next((i for i, v in enumerate(row_cleaned) if any(k in v for k in kw_ma)), -1)
                            tmp_ten = next((i for i, v in enumerate(row_cleaned) if any(k in v for k in kw_ten)), -1)
                            
                            if tmp_ma != -1 and tmp_ten != -1:
                                c_ma, c_ten = tmp_ma, tmp_ten
                                continue
                                
                            if c_ma != -1 and c_ten != -1:
                                m_val = row_raw[c_ma].replace('.0', '').strip()
                                t_val = row_raw[c_ten].strip()
                                if m_val and m_val.lower() not in ['none', 'nan', '']:
                                    if m_val not in dict_nv_chuan: dict_nv_chuan[m_val] = t_val

                    # 2. Xử lý File Học viên
                    results = []
                    total_checked = 0
                    dict_hv_sheets = read_excel_values_only(file_dshv)
                    
                    for sheet_name, df_sheet in dict_hv_sheets.items():
                        c_ma, c_ten = -1, -1
                        current_class = "N/A"
                        pending_class_name = ""
                        for idx, row in df_sheet.iterrows():
                            row_raw = [str(x).strip() if x is not None else "" for x in row.values]
                            row_cleaned = [clean_header(x) for x in row_raw]
                            
                            cell_0 = row_raw[0] 
                            cell_6 = row_raw[6] if len(row_raw) > 6 else "" 
                            cell_7 = row_raw[7] if len(row_raw) > 7 else "" 
                            
                            if cell_0 and "/" in str(cell_6):
                                current_class = f"{cell_0.split(chr(10))[0]} [{cell_6} - {cell_7}]"
                                pending_class_name = "" 
                            elif cell_0 and len(cell_0) > 10 and not any(k in clean_header(cell_0) for k in kw_ma + kw_ten):
                                pending_class_name = cell_0.split(chr(10))[0]
                                
                            if ("lý thuyết" in cell_0.lower() or "thực hành" in cell_0.lower()) and "/" in str(cell_6):
                                if pending_class_name:
                                    current_class = f"{pending_class_name} [{cell_6} - {cell_7}]"

                            tmp_ma = next((i for i, v in enumerate(row_cleaned) if any(k in v for k in kw_ma)), -1)
                            tmp_ten = next((i for i, v in enumerate(row_cleaned) if any(k in v for k in kw_ten)), -1)
                            
                            if tmp_ma != -1 and tmp_ten != -1:
                                c_ma, c_ten = tmp_ma, tmp_ten
                                continue
                                
                            if c_ma != -1 and c_ten != -1:
                                ma_nv = row_raw[c_ma].replace('.0', '').strip()
                                if ma_nv and ma_nv.lower() not in ['none', 'nan', '']:
                                    ho_ten = row_raw[c_ten]
                                    total_checked += 1
                                    
                                    if ma_nv in dict_nv_chuan:
                                        ten_chuan = dict_nv_chuan[ma_nv]
                                        if " ".join(ho_ten.lower().split()) != " ".join(ten_chuan.lower().split()):
                                            results.append({
                                                "Lớp/Khóa": current_class, "Trang": sheet_name, "Mã NV": ma_nv,
                                                "Tên trong file": ho_ten, "Tên đúng (DSNV)": ten_chuan, "Lỗi": "Sai họ tên"
                                            })
                                    else:
                                        results.append({
                                            "Lớp/Khóa": current_class, "Trang": sheet_name, "Mã NV": ma_nv,
                                            "Tên trong file": ho_ten, "Tên đúng (DSNV)": "❌ KHÔNG CÓ TRONG GỐC", "Lỗi": "Mã NV sai/lạ"
                                        })

                    # 3. Hiển thị báo cáo
                    st.divider()
                    st.subheader(f"📊 Kết quả kiểm tra (Tổng quét: {total_checked} HV)")
                    if not results:
                        st.success("🎉 Tuyệt vời! Danh sách khớp 100% với dữ liệu gốc.")
                    else:
                        df_res = pd.DataFrame(results)
                        st.error(f"Phát hiện {len(df_res)} lỗi cần chỉnh sửa.")
                        st.dataframe(df_res, use_container_width=True, hide_index=True)
                        
                        csv = df_res.to_csv(index=False).encode('utf-8-sig')
                        st.download_button("📥 Tải danh sách lỗi (.csv)", data=csv, file_name='loi_danh_sach.csv', mime='text/csv')

                except Exception as e:
                    st.error(f"❌ Lỗi xử lý: {str(e)}")

# --- TAB 2: CHỨC NĂNG TẠO KHUNG TUẦN TỪ KHĐT ---
with tab_tao_khung:
    st.info("💡 Bước 1: Upload KHĐT để tạo ra file Excel chuẩn (Có mục lục, khung rỗng sẵn sàng đón học viên).")
    file_khdt = st.file_uploader("📂 Chọn file Kế hoạch đào tạo (KHĐT)", type=["xlsx"])
    
    if file_khdt:
        if "dict_khdt" not in st.session_state or st.session_state.get("file_khdt_name") != file_khdt.name:
            with st.spinner("Đang phân tích dữ liệu KHĐT..."):
                st.session_state["dict_khdt"] = doc_khdt_gom_theo_tuan(file_khdt)
                st.session_state["file_khdt_name"] = file_khdt.name
        
        dict_khdt = st.session_state["dict_khdt"]
        danh_sach_tuan = list(dict_khdt.keys())
        
        chon_tat_ca = st.checkbox("Tạo toàn bộ các tuần (Xuất file ZIP)")
        tuan_duoc_chon = danh_sach_tuan if chon_tat_ca else st.multiselect("Hoặc tuỳ chọn các tuần cụ thể (Nếu chỉ chọn 1 tuần, sẽ tải thẳng file Excel):", options=danh_sach_tuan, default=danh_sach_tuan[:1] if danh_sach_tuan else [])
            
        if st.button("🚀 Tạo Template rỗng", type="primary"):
            if tuan_duoc_chon:
                dict_chon = {k: v for k, v in dict_khdt.items() if k in tuan_duoc_chon}
                if len(dict_chon) == 1:
                    tuan_don = list(dict_chon.keys())[0]
                    excel_file = tao_file_excel_mot_tuan(tuan_don, dict_chon[tuan_don])
                    safe_tuan = str(tuan_don).replace("/", "-").replace(":", "")
                    st.download_button("📥 Tải Template Excel", data=excel_file, file_name=f"Template_DSLop_{safe_tuan}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    zip_file = tao_file_excel_tu_dict(dict_chon)
                    st.download_button("📥 Tải Template ZIP", data=zip_file, file_name=f"Template_DSLop_Full.zip", mime="application/zip")
            else:
                st.warning("⚠️ Vui lòng chọn ít nhất 1 tuần!")

# --- TAB 3: CHỨC NĂNG NHỒI HỌC VIÊN TỰ ĐỘNG ---
with tab_nhoi_hv:
    st.info("💡 Tính năng đọc thông tin lớp ở ô D7 và D9 của file Khung (Bước 2) để tự động bốc toàn bộ Học viên ở file Danh sách tháng (DSHV) dán vào trang tương ứng.")
    
    col_x, col_y = st.columns(2)
    with col_x:
        file_template_in = st.file_uploader("📂 1. Chọn file Khung rỗng (File tuần đã tạo ở Bước 2)", type=["xlsx"], key="tpl_in")
    with col_y:
        file_dshv_in = st.file_uploader("📂 2. Chọn file Danh sách Học viên tổng (VD: File DS T6)", type=["xlsx"], key="dshv_in")
        
    if st.button("🪄 Bắt đầu khớp & Điền học viên", key="btn_nhoi", type="primary"):
        if file_template_in and file_dshv_in:
            with st.spinner("Đang dò tìm chéo dữ liệu Lớp và Học viên..."):
                ds_lop_hv = doc_dshv_ra_list(file_dshv_in)
                filled_excel = nhoi_hoc_vien_vao_template(file_template_in, ds_lop_hv)
                
                st.success("🎉 Khớp dữ liệu thành công!")
                st.download_button(
                    label="📥 Tải File Danh Sách Lớp Hoàn Chỉnh",
                    data=filled_excel,
                    file_name=file_template_in.name.replace("Template_", "DS_Lop_ChinhThuc_"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ Bạn cần tải lên đủ cả file Khung Tuần và file Danh Sách Học Viên tháng để thực hiện!")