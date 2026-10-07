import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font
import io
import zipfile
import re
import os
import shutil
import sqlite3
from datetime import datetime

from openpyxl.cell.cell import MergedCell

import khdt_core
import khdt_export
import khdt_store
from khdt_core import LOAI_HINH, LOAI_HINH_FILE, MIEN, MIEN_FILE, bo_dau


def ten_muc(ma):
    """Tên hiển thị của một mục xuất (6 loại hình + 'DOI_TAC')."""
    return "Đối tác" if ma == "DOI_TAC" else LOAI_HINH[ma]

# ==========================================
# CẤU HÌNH GIAO DIỆN & KHỞI TẠO CSDL
# ==========================================
st.set_page_config(page_title="🛠️ Công cụ Xử lý Dữ liệu Đào tạo VIAGS", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    .status-box { padding: 20px; border-radius: 10px; border: 1px solid #dee2e6; background-color: white; }
    </style>
""", unsafe_allow_html=True)

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dsnv_local.db")

def get_db_connection():
    """Tạo kết nối tới SQLite và trả về dạng Row để truy cập bằng tên cột"""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Khởi tạo cấu trúc bảng CSDL - Chuẩn cấu trúc để đồng bộ Supabase sau này"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS nhan_vien (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ma_nv TEXT,
            ho_ten TEXT,
            don_vi_goc TEXT,
            don_vi_chuan TEXT
        )
    """)
    # Tạo Index cho ma_nv để tăng tốc độ truy vấn tối đa khi danh sách phình to
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_ma_nv ON nhan_vien (ma_nv)")
    conn.commit()
    khdt_store.khoi_tao_bang(conn)  # bảng khdt / lop_hoc / lop_thang
    conn.close()

# Chạy khởi tạo bảng ngay khi ứng dụng kích hoạt
init_db()

# ==========================================
# 📑 HÀM TIỆN ÍCH DÙNG CHUNG
# ==========================================
def remove_vietnamese_accents(s):
    return bo_dau(s)

def chuan_ma_nv(v):
    """Mã NV dạng chuỗi: bỏ khoảng trắng và đuôi '.0' do Excel đọc số (không đụng '.0' ở giữa mã)."""
    return re.sub(r'\.0$', '', str(v).strip())

def tao_sao_luu_db():
    """Sao lưu CSDL trước khi ghi đè; trả về đường dẫn file sao lưu (hoặc None nếu chưa có CSDL)."""
    if not os.path.exists(DB_FILE):
        return None
    thu_muc = os.path.join(os.path.dirname(DB_FILE), "backup")
    os.makedirs(thu_muc, exist_ok=True)
    dich = os.path.join(thu_muc, f"dsnv_local_{datetime.now():%Y%m%d_%H%M%S}.db")
    shutil.copy2(DB_FILE, dich)
    return dich

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

def chuan_hoa_don_vi(dv_str):
    if not dv_str or str(dv_str).lower() == 'nan': return ""
    dv_clean = re.sub(r'[\s\.\-\_]', '', str(dv_str).lower())
    dv_no_accent = remove_vietnamese_accents(dv_clean).lower() 
    
    mapping = {
        "hanhkhach": "PVHK", "pvhk": "PVHK", "hk": "PVHK",
        "hanghoa": "PVHH", "pvhh": "PVHH", "hh": "PVHH",
        "dieuhanh": "TTĐH", "ttdh": "TTĐH", "dh": "TTĐH", "đh": "TTĐH",
        "sando": "PVSĐ", "pvsd": "PVSĐ", "sd": "PVSĐ", "sđ": "PVSĐ",
        "hanhchinh": "KHHC", "khhc": "KHHC", "hc": "KHHC",
        "trung tâm huấn luyện": "VNBA", "Trung tâm Huấn luyện": "VNBA", "trungtamhuanluyen": "VNBA", "vnba": "VNBA",
        "ketoan": "KTOA", "ktoa": "KTOA", "kt": "KTOA",
        "chatluong": "QLCL", "qlcl": "QLCL", "cl": "QLCL"
    }
    for key, val in mapping.items():
        if key in dv_no_accent:
            return val
    return str(dv_str).strip().upper()

# ==========================================
# CORE LOGIC BÓC TÁCH KHĐT & TRỘN HỌC VIÊN
# ==========================================

def doc_dshv_ra_list(file_dshv):
    file_dshv.seek(0)
    wb_dshv = pd.ExcelFile(file_dshv)
    ds_lop_hv = []
    
    for sheetname in wb_dshv.sheet_names:
        s_name_clean = remove_vietnamese_accents(sheetname).lower()
        if s_name_clean in ["mucluc", "sheet1"]: continue
        
        if any(k in s_name_clean for k in ['ban dau', 'bandau', 'bsn', 'bsnd']): loai_hinh_sheet = 'Ban đầu'
        elif 'dinh ky' in s_name_clean or 'dinhky' in s_name_clean: loai_hinh_sheet = 'Định kỳ'
        elif 'phuc hoi' in s_name_clean or 'phuchoi' in s_name_clean: loai_hinh_sheet = 'Phục hồi'
        else: loai_hinh_sheet = 'Bồi dưỡng kiến thức' 
            
        df_sheet = pd.read_excel(wb_dshv, sheet_name=sheetname, header=None)
        header_row_hv = next((idx for idx, row in df_sheet.iterrows() if any("khóa học" in str(s).lower() or "khoa hoc" in str(s).lower() for s in row.values)), None)
        if header_row_hv is None: continue
        current_hv_class = None
        
        for idx in range(header_row_hv + 1, len(df_sheet)):
            row = df_sheet.iloc[idx]
            val_khoa_hoc = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            
            if val_khoa_hoc != "" and val_khoa_hoc != "nan" and not val_khoa_hoc.startswith("C."):
                tu_ngay_hv = str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else ""
                den_ngay_hv = str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else ""
                
                if tu_ngay_hv == "nan": tu_ngay_hv = ""
                if den_ngay_hv == "nan": den_ngay_hv = ""
                
                tg_hv = f"{tu_ngay_hv} - {den_ngay_hv}" if tu_ngay_hv != den_ngay_hv and den_ngay_hv else tu_ngay_hv
                ten_lop_hv = val_khoa_hoc.split("\n")[0].strip()
                khoang_hv = khdt_core.khoang_ngay(row.iloc[6], row.iloc[7])

                current_hv_class = {"ten_lop": ten_lop_hv, "thoi_gian": tg_hv, "khoang": khoang_hv, "loai_hinh": loai_hinh_sheet, "hoc_vien": []}
                ds_lop_hv.append(current_hv_class)
                
            elif current_hv_class is not None:
                ma_nv = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
                ho_ten = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
                don_vi = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ""
                if ho_ten and ho_ten != "nan" and ho_ten != "HỌ VÀ TÊN":
                    current_hv_class["hoc_vien"].append({
                        "manv": ma_nv if ma_nv != "nan" else "", 
                        "hoten": ho_ten, 
                        "donvi": don_vi if don_vi != "nan" else ""
                    })
    return ds_lop_hv

def nhoi_hoc_vien_vao_template(file_template, ds_lop_hv):
    def norm_name(s):
        s = remove_vietnamese_accents(str(s)).lower()
        s = re.sub(r'[^a-z0-9]', ' ', s)
        return " ".join(s.split()) + " "

    def tim_bang_hoc_vien(ws):
        """Vị trí bảng học viên trong sheet lớp: (dòng đầu dữ liệu, cột 'Mã NV') hoặc None.
        Mẫu TTĐT: tiêu đề gộp 2 dòng (12-13), dữ liệu từ dòng 14, STT ở cột B (công thức) → Mã NV ở cột C."""
        for r in range(8, 16):
            for c in range(1, 8):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and remove_vietnamese_accents(v).lower().replace(" ", "") == "manv":
                    bat_dau = r + 1
                    while isinstance(ws.cell(row=bat_dau, column=c), MergedCell):
                        bat_dau += 1
                    return bat_dau, c
        return None

    file_template.seek(0)
    wb = openpyxl.load_workbook(file_template)

    for sheetname in wb.sheetnames:
        if remove_vietnamese_accents(sheetname).lower().replace(" ", "") == "mucluc": continue
        ws = wb[sheetname]

        ten_lop_kh = str(ws['D7'].value or "")
        tg_kh = str(ws['D9'].value or "")
        b8_text = str(ws['B8'].value or "")

        loai_hinh_kh = ""
        m_loai = re.search(r'đào tạo:\s*([^/]+)', b8_text, re.IGNORECASE) or re.search(r'Loại hình:\s*([^/]+)', b8_text)
        if m_loai: loai_hinh_kh = m_loai.group(1).strip()

        key_kh_ten = norm_name(ten_lop_kh)
        if not key_kh_ten.strip():
            continue
        kh_khoang = khdt_core.khoang_ngay(tg_kh)
        key_kh_loai = norm_name(loai_hinh_kh)[:6]

        ung_vien = []
        for hv_class in ds_lop_hv:
            key_hv_ten = norm_name(hv_class["ten_lop"])
            if not key_hv_ten.strip():
                continue
            hv_khoang = hv_class.get("khoang")
            key_hv_loai = norm_name(hv_class["loai_hinh"])[:6]

            name_match = (key_kh_ten in key_hv_ten) or (key_hv_ten in key_kh_ten)
            type_match = (key_kh_loai == key_hv_loai)
            date_match = bool(kh_khoang and hv_khoang and kh_khoang[0] <= hv_khoang[1] and hv_khoang[0] <= kh_khoang[1])
            ten_giong_het = key_kh_ten == key_hv_ten

            if name_match and type_match and (date_match or ten_giong_het):
                ung_vien.append(((ten_giong_het, date_match), hv_class))

        # chọn lớp khớp nhất (trùng tên hoàn toàn + trùng ngày), không lấy bừa lớp đầu tiên
        matched_hv = max(ung_vien, key=lambda x: x[0])[1]["hoc_vien"] if ung_vien else []
        vi_tri = tim_bang_hoc_vien(ws) if matched_hv else None

        if matched_hv and vi_tri:
            bat_dau, c_ma = vi_tri
            for r in range(bat_dau, bat_dau + max(len(matched_hv), 15)):
                for c in range(c_ma, c_ma + 3): ws.cell(row=r, column=c).value = None
            for i, hv in enumerate(matched_hv):
                r_idx = bat_dau + i
                if ws.cell(row=r_idx, column=c_ma - 1).value is None:  # mẫu đã có công thức STT thì giữ
                    ws.cell(row=r_idx, column=c_ma - 1, value=i + 1)
                ws.cell(row=r_idx, column=c_ma, value=hv["manv"])
                ws.cell(row=r_idx, column=c_ma + 1, value=hv["hoten"])
                ws.cell(row=r_idx, column=c_ma + 2, value=hv["donvi"])

    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# ==========================================
# GIAO DIỆN CHÍNH (STREAMLIT TABS)
# ==========================================
st.title("🚀 Hệ thống Trộn & Đối chiếu Dữ liệu Đào tạo VIAGS")

tab_doi_chieu, tab_tao_khung, tab_nhoi_hv, tab_ql_csdl = st.tabs([
    "🔍 1. ĐỐI CHIẾU DANH SÁCH HỌC VIÊN", 
    "📄 2. TẠO KHUNG TỪ KHĐT", 
    "🧑‍🎓 3. TỰ ĐỘNG THÊM HỌC VIÊN",
    "🗄️ 4. QUẢN LÝ CƠ SỞ DỮ LIỆU"
])

# --- TAB 1: ĐỐI CHIẾU DANH SÁCH (TRUY VẤN REAL-TIME TỪ SQLITE) ---
with tab_doi_chieu:
    st.info("💡 Hướng dẫn: Tải lên file Học viên cần rà soát. Hệ thống tự động quét tìm chéo với Cơ sở dữ liệu nội bộ.")
    file_dshv = st.file_uploader("Chọn file Danh sách học viên cần kiểm tra", type=["xlsx", "xls", "csv"])

    if st.button("🚀 Bắt đầu quét đối chiếu", type="primary"):
        if not file_dshv:
            st.warning("⚠️ Vui lòng tải lên File Danh sách học viên cần kiểm tra!")
        else:
            with st.spinner("Đang truy vấn Cơ sở dữ liệu và phân tích dữ liệu..."):
                try:
                    # Kiểm tra CSDL có dữ liệu không trước khi chạy
                    conn = get_db_connection()
                    count_nv = conn.execute("SELECT COUNT(*) FROM nhan_vien").fetchone()[0]
                    if count_nv == 0:
                        st.error("⚠️ Cơ sở dữ liệu hiện đang TRỐNG! Vui lòng qua Tab 4 nạp file DSNV master trước.")
                        conn.close()
                        st.stop()

                    results = []
                    total_checked = 0
                    dict_hv_sheets = read_excel_values_only(file_dshv)
                    
                    kw_ma = ['mãnv', 'mnv', 'manv', 'mãsốnv', 'mãnhânviên', 'staffid']
                    kw_ten = ['họvàtên', 'họtên', 'fullname']
                    
                    for sheet_name, df_sheet in dict_hv_sheets.items():
                        c_ma, c_ten, c_dv = -1, -1, -1
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
                            tmp_dv = next((i for i, v in enumerate(row_cleaned) if ('trungtâm' in v or 'trungtam' in v) and 'đội' not in v and 'doi' not in v), -1)
                            
                            if tmp_ma != -1 and tmp_ten != -1:
                                c_ma, c_ten = tmp_ma, tmp_ten
                                if tmp_dv != -1: c_dv = tmp_dv
                                continue
                                
                            if c_ma != -1 and c_ten != -1:
                                ma_nv = chuan_ma_nv(row_raw[c_ma])
                                if ma_nv and ma_nv.lower() not in ['none', 'nan', '']:
                                    ho_ten = row_raw[c_ten]
                                    total_checked += 1
                                    
                                    dv_hv_raw = row_raw[c_dv] if c_dv != -1 else (row_raw[5] if len(row_raw) > 5 else "")
                                    dv_hv_chuan = chuan_hoa_don_vi(dv_hv_raw)
                                    
                                    # --- TRUY VẤN REAL-TIME TỪ SQLITE ---
                                    # Lấy toàn bộ danh sách nhân viên khớp mã này (giải quyết triệt để trùng mã đa nhánh)
                                    cursor = conn.cursor()
                                    cursor.execute("SELECT ho_ten, don_vi_chuan FROM nhan_vien WHERE ma_nv = ?", (ma_nv,))
                                    candidates = cursor.fetchall()
                                    
                                    if candidates:
                                        best_match = None
                                        loi_msg_best = ["Sai họ tên", "Sai đơn vị"]
                                        
                                        for row_nv in candidates:
                                            ten_chuan = row_nv["ho_ten"]
                                            dv_chuan = row_nv["don_vi_chuan"]
                                            
                                            sai_ten = " ".join(ho_ten.lower().split()) != " ".join(ten_chuan.lower().split())
                                            sai_dv = (dv_chuan != "") and (dv_hv_chuan != dv_chuan)
                                            
                                            if not sai_ten and not sai_dv:
                                                loi_msg_best = []
                                                best_match = row_nv
                                                break
                                                
                                            current_loi = []
                                            if sai_ten: current_loi.append("Sai họ tên")
                                            if sai_dv: current_loi.append("Sai đơn vị")
                                            
                                            if len(current_loi) < len(loi_msg_best):
                                                loi_msg_best = current_loi
                                                best_match = row_nv
                                                
                                        if loi_msg_best:
                                            if best_match is None: best_match = candidates[0]
                                            results.append({
                                                "Lớp/Khóa": current_class, "Trang": sheet_name, "Mã NV": ma_nv,
                                                "Tên (File)": ho_ten, 
                                                "Tên đúng": best_match["ho_ten"] if "Sai họ tên" in loi_msg_best else "-",
                                                "Đơn vị (File)": dv_hv_chuan if "Sai đơn vị" in loi_msg_best else "-",
                                                "Đơn vị đúng": best_match["don_vi_chuan"] if "Sai đơn vị" in loi_msg_best else "-",
                                                "Lỗi": " & ".join(loi_msg_best)
                                            })
                                    else:
                                        results.append({
                                            "Lớp/Khóa": current_class, "Trang": sheet_name, "Mã NV": ma_nv,
                                            "Tên (File)": ho_ten, "Tên đúng": "❌ KHÔNG CÓ TRONG GỐC",
                                            "Đơn vị (File)": dv_hv_chuan, "Đơn vị đúng": "-", "Lỗi": "Mã NV lạ"
                                        })
                    conn.close()

                    st.divider()
                    st.subheader(f"📊 Kết quả kiểm tra (Tổng quét: {total_checked} HV)")
                    if not results:
                        st.success("🎉 Tuyệt vời! Danh sách khớp 100% với Cơ sở dữ liệu gốc.")
                    else:
                        df_res = pd.DataFrame(results)
                        st.error(f"Phát hiện {len(df_res)} lỗi cần chỉnh sửa.")
                        st.dataframe(df_res, use_container_width=True, hide_index=True)
                        
                        csv = df_res.to_csv(index=False).encode('utf-8-sig')
                        st.download_button("📥 Tải danh sách lỗi (.csv)", data=csv, file_name='loi_danh_sach.csv', mime='text/csv')

                except Exception as e:
                    st.error(f"❌ Lỗi xử lý dữ liệu: {str(e)}")

# --- TAB 2: TẠO DANH SÁCH LỚP KHUNG TỪ KHĐT (MỖI LẦN MỘT MIỀN) ---
def _dat_nhom_checkbox(prefix, keys, gia_tri):
    for k in keys:
        st.session_state[f"{prefix}_{k}"] = gia_tri


def checkbox_nhom(prefix, tuy_chon, mac_dinh, so_cot=3):
    """Nhóm checkbox + nút Chọn tất cả / Bỏ chọn. tuy_chon: {mã: nhãn}. Trả về list mã được tích."""
    for k in tuy_chon:
        st.session_state.setdefault(f"{prefix}_{k}", k in mac_dinh)
    b1, b2, _ = st.columns([1, 1, 4])
    b1.button("Chọn tất cả", key=f"{prefix}_btn_all", on_click=_dat_nhom_checkbox, args=(prefix, list(tuy_chon), True))
    b2.button("Bỏ chọn", key=f"{prefix}_btn_none", on_click=_dat_nhom_checkbox, args=(prefix, list(tuy_chon), False))
    cot = st.columns(so_cot)
    chon = []
    for i, (k, nhan) in enumerate(tuy_chon.items()):
        if cot[i % so_cot].checkbox(nhan, key=f"{prefix}_{k}"):
            chon.append(k)
    return chon


def doc_mau_theo_loai(ma_loai):
    """(bytes file mẫu | None, mô tả nguồn mẫu). Ưu tiên file người dùng tải lên, rồi file theo mau_ds_lop.json."""
    up = st.session_state.get(f"mau_up_{ma_loai}")
    if up is not None:
        return up.getvalue(), f"mẫu tải lên: {up.name}"
    duong_dan = khdt_export.tim_file_mau(ma_loai)
    if duong_dan:
        try:
            with open(duong_dan, "rb") as f:
                return f.read(), f"mẫu: {os.path.basename(duong_dan)}"
        except OSError as e:
            return None, f"không đọc được mẫu ({os.path.basename(duong_dan)}: {e}) → dùng khung cơ bản"
    return None, "không có file mẫu → dùng khung cơ bản"


with tab_tao_khung:
    st.info("💡 Chọn MỘT miền → upload KHĐT → phân tích (tự lưu vào CSDL) → xuất danh sách lớp khung của miền đó theo tháng của KHĐT.")

    # ---------- BƯỚC 1: CHỌN MIỀN + ĐỌC KHĐT (TỰ LƯU) ----------
    st.subheader("Bước 1. Chọn miền và đọc KHĐT")
    mien_chon = st.radio("Miền cần tạo", list(MIEN), format_func=MIEN.get, horizontal=True, key="p1_mien")
    file_khdt = st.file_uploader("📂 Chọn file Kế hoạch đào tạo (KHĐT)", type=["xlsx"], key="khdt_up")

    if file_khdt:
        du_lieu_khdt = file_khdt.getvalue()
        khoa_file = (file_khdt.name, len(du_lieu_khdt))
        if st.session_state.get("khdt_khoa") != khoa_file:
            try:
                with st.spinner("Đang đọc file KHĐT..."):
                    wb_k = openpyxl.load_workbook(io.BytesIO(du_lieu_khdt), data_only=True, read_only=True)
                    st.session_state["khdt_rows"] = {w.title: [list(r) for r in w.iter_rows(values_only=True)] for w in wb_k.worksheets}
                st.session_state["khdt_khoa"] = khoa_file
                st.session_state.pop("khdt_kq", None)
                st.session_state.pop("khung_out", None)
            except Exception as e:
                st.session_state.pop("khdt_khoa", None)
                st.error(f"❌ Không đọc được file KHĐT (cần file .xlsx hợp lệ): {e}")

    rows_by_sheet = st.session_state.get("khdt_rows") if file_khdt and st.session_state.get("khdt_khoa") else None

    if rows_by_sheet:
        ten_sheets = list(rows_by_sheet)
        thang_goi_y, nam_goi_y = next(((t, n) for t, n in (khdt_core.doan_thang_nam(r) for r in rows_by_sheet.values()) if t), (None, None))
        goi_y = khdt_core.goi_y_sheet(ten_sheets, mien_chon)
        if not goi_y:
            st.warning(f"⚠️ Không tìm thấy sheet gợi ý cho {MIEN[mien_chon]} — hãy chọn đúng sheet bên dưới.")
        c_s, c_t, c_n = st.columns([3, 1, 1])
        sheet_chon = c_s.selectbox(f"Sheet dữ liệu của {MIEN[mien_chon]}", ten_sheets,
                                   index=ten_sheets.index(goi_y[0]) if goi_y else 0, key=f"p1_sheet_{mien_chon}")
        thang_khdt = c_t.number_input("Tháng KHĐT", 1, 12, thang_goi_y or datetime.now().month)
        nam_khdt = c_n.number_input("Năm KHĐT", 2000, 2100, nam_goi_y or datetime.now().year)

        if st.button("🔎 Phân tích và lưu KHĐT", type="primary"):
            try:
                with st.spinner("Đang phân tích và lưu..."):
                    lops, cb = khdt_core.doc_khdt_sheet(rows_by_sheet[sheet_chon], mien_chon, int(thang_khdt), int(nam_khdt))
                    conn_k = get_db_connection()
                    try:
                        khdt_store.luu_khdt(conn_k, int(thang_khdt), int(nam_khdt), mien_chon, sheet_chon, file_khdt.name, lops)
                    finally:
                        conn_k.close()
                st.session_state["khdt_kq"] = dict(thang=int(thang_khdt), nam=int(nam_khdt), mien=mien_chon, sheet=sheet_chon,
                                                   ten_file=file_khdt.name, lops=lops, canh_bao=cb)
                st.session_state.pop("khung_out", None)
            except Exception as e:
                st.error(f"❌ Không phân tích/lưu được: {e}")

    kq = st.session_state.get("khdt_kq")
    if kq and kq["mien"] == mien_chon:
        m, thang_x, nam_x, lops = kq["mien"], kq["thang"], kq["nam"], kq["lops"]
        st.success(f"✅ Đã đọc và lưu KHĐT T{thang_x:02d}/{nam_x} — {MIEN[m]} (sheet '{kq['sheet']}'): {len(lops)} lớp "
                   "(KHĐT cùng tháng/năm/miền được ghi đè).")
        dem = {LOAI_HINH[k]: sum(1 for x in lops if x["loai_hinh"] == k) for k in LOAI_HINH}
        for c, (ten, n) in zip(st.columns(len(dem)), dem.items()):
            c.metric(ten, n)
        if m == "VNA_DT":
            dt = {}
            for x in lops:
                if x.get("doi_tac"):
                    dt[x["doi_tac"]] = dt.get(x["doi_tac"], 0) + 1
            st.caption(f"VNA: {sum(1 for x in lops if not x.get('doi_tac'))} lớp | Đối tác: "
                       + (", ".join(f"{t} ({n} lớp)" for t, n in dt.items()) or "không có"))
        if kq["canh_bao"]:
            with st.expander(f"⚠️ {len(kq['canh_bao'])} cảnh báo (dòng bị bỏ qua / không đọc được)"):
                st.write("\n".join(f"- {w}" for w in kq["canh_bao"]))
        with st.expander("Xem danh sách lớp đã đọc"):
            st.dataframe(pd.DataFrame([{
                "Loại hình": LOAI_HINH[x["loai_hinh"]], "Đối tác": x.get("doi_tac") or "", "Tên lớp": x["ten_lop"],
                "Thời gian": x["thoi_gian"], "Hình thức": x["hinh_thuc"], "Địa điểm": x["dia_diem"],
                "Giáo viên": x["giao_vien"], "SL HV": x["so_hv"], "Dòng KHĐT": x["dong_goc"]} for x in lops]),
                use_container_width=True, hide_index=True)

        # ---------- BƯỚC 2: XUẤT (CHỈ MIỀN VỪA CHỌN, THÁNG CỦA KHĐT) ----------
        st.divider()
        st.subheader(f"Bước 2. Xuất danh sách lớp khung — {MIEN[m]} (tháng {thang_x:02d}/{nam_x})")
        muc_chon, xuat_vna, xuat_dt = None, True, True
        mau_keys = dict(LOAI_HINH)
        if m == "VNA_DT":
            lua_chon = st.radio("Nội dung xuất", ["VNA (mỗi loại hình một file)", "Đối tác (mỗi đối tác 1 file)", "Cả VNA và đối tác"],
                                horizontal=True)
            xuat_vna, xuat_dt = not lua_chon.startswith("Đối tác"), not lua_chon.startswith("VNA (")
            if xuat_dt:
                mau_keys["DOI_TAC"] = "Đối tác"
        if xuat_vna:
            st.markdown("**Mục (loại hình) cần xuất** — mỗi mục một file"
                        + ("; Ban đầu – NVM tự tách mỗi khóa một file:" if m != "VNA_DT" else ":"))
            muc_chon = checkbox_nhom("p2_loai", LOAI_HINH, list(LOAI_HINH), so_cot=3)
        che_do_nhan = st.radio("Sắp xếp lớp (thứ tự sheet) trong mỗi file",
                               ["Theo KHĐT (đúng thứ tự trong KHĐT)", "Theo thời gian tăng dần"], horizontal=True)
        che_do = "khdt" if che_do_nhan.startswith("Theo KHĐT") else "thoi_gian"
        kieu_xuat = st.radio(
            "Kiểu file xuất",
            ["Điền vào file mẫu (đủ logo/định dạng, mỗi file vài giây)",
             "File nhanh: chỉ có dữ liệu, đúng vị trí ô như mẫu (tự copy vào file mẫu)"],
            index=1 if khdt_export.TREN_CLOUD else 0,  # máy chủ cloud ít RAM/CPU → mặc định file nhanh
            horizontal=False)
        xuat_nhanh = kieu_xuat.startswith("File nhanh")

        with st.expander("File mẫu danh sách lớp (không bắt buộc)"):
            st.caption("Mặc định lấy mẫu theo cấu hình mau_ds_lop.json; tải file ở đây để dùng mẫu khác cho mục tương ứng.")
            for k, ten in mau_keys.items():
                st.file_uploader(f"Mẫu cho: {ten}", type=["xlsx"], key=f"mau_up_{k}")

        if st.button("📄 Tạo danh sách lớp khung", type="primary", disabled=(xuat_vna and not muc_chon and not (m == "VNA_DT" and xuat_dt))):
            conn_x = get_db_connection()
            try:
                lops_m = khdt_store.doc_lop(conn_x, m, nam_x, thang_x, None, che_do)
            finally:
                conn_x.close()
            cac_file = khdt_core.chia_file_xuat(lops_m, m, muc_chon, xuat_vna, xuat_dt)
            dau_ra, viec = [], []
            with st.spinner("Đang tạo file (mỗi file nạp mẫu vài giây, các file được tạo song song)..."):
                mau_da_doc = {f["loai"]: ((None, "file nhanh: chỉ có dữ liệu") if xuat_nhanh else doc_mau_theo_loai(f["loai"]))
                              for f in cac_file}
                for f in cac_file:
                    mau_bytes, nguon = mau_da_doc[f["loai"]]
                    viec.append(dict(k=f["loai"], hau_to=f["hau_to"], lops=f["lops"], mau=mau_bytes, nguon=nguon))
                # file có mẫu: tạo song song; không có mẫu: dựng khung cơ bản (nhanh)
                co_mau = [v for v in viec if v["mau"]]
                for v, kq_v in zip(co_mau, khdt_export.dien_nhieu_mau([(v["mau"], v["lops"], v["k"]) for v in co_mau])):
                    v["ket_qua"] = kq_v
                for v in viec:
                    if not v["mau"]:
                        v["ket_qua"] = khdt_export.tao_khung_co_ban(v["lops"], v["k"])
                for v in viec:
                    if isinstance(v["ket_qua"], Exception):
                        st.error(f"❌ {ten_muc(v['k'])} {v['hau_to']}: {v['ket_qua']}")
                        continue
                    if v["k"] == "DOI_TAC":
                        ten_f = f"DS lop hoc DOI TAC {v['hau_to']} T{thang_x:02d}-{nam_x}"
                    else:
                        hau_to = f" {v['hau_to']}" if v["hau_to"] else ""
                        ten_f = f"DS lop hoc {MIEN_FILE[m]} {LOAI_HINH_FILE[v['k']]}{hau_to} T{thang_x:02d}-{nam_x}"
                    dau_ra.append(dict(mien=m, loai=v["k"], hau_to=v["hau_to"], ten=khdt_export.ten_file_an_toan(ten_f) + ".xlsx",
                                       data=v["ket_qua"], so_lop=len(v["lops"]), nguon=v["nguon"]))
            bo_qua = [ten_muc(k) for k in (muc_chon or []) if not any(f["loai"] == k for f in cac_file)]
            st.session_state["khung_out"] = dict(files=dau_ra, bo_qua=bo_qua, thang=thang_x, nam=nam_x)

        out = st.session_state.get("khung_out")
        if out:
            if out["bo_qua"]:
                st.caption("Không có lớp (không tạo file): " + "; ".join(out["bo_qua"]))
            if not out["files"]:
                st.warning("Không có file nào được tạo.")
            else:
                st.success(f"🎉 Đã tạo {len(out['files'])} file.")
                st.dataframe(pd.DataFrame([{"Mục": ten_muc(f["loai"]), "Khóa / Đối tác": f["hau_to"], "Số lớp (sheet)": f["so_lop"],
                                            "File": f["ten"], "Mẫu": f["nguon"]} for f in out["files"]]),
                             use_container_width=True, hide_index=True)
                if len(out["files"]) > 1:
                    zbuf = io.BytesIO()
                    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for f in out["files"]:
                            zf.writestr(f["ten"], f["data"])
                    st.download_button("📦 Tải tất cả (ZIP)", zbuf.getvalue(),
                                       file_name=f"DS_lop_khung_{MIEN_FILE[m]}_T{out['thang']:02d}-{out['nam']}.zip",
                                       mime="application/zip", type="primary")
                for i, f in enumerate(out["files"]):
                    nhan_f = ten_muc(f["loai"]) + (f" {f['hau_to']}" if f["hau_to"] else "")
                    st.download_button(f"📥 {nhan_f} ({f['so_lop']} lớp)", f["data"], file_name=f["ten"], key=f"dl_khung_{i}",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- TAB 3: CHỨC NĂNG NHỒI HỌC VIÊN TỰ ĐỘNG ---
with tab_nhoi_hv:
    st.info("💡 Tính năng đọc thông tin lớp ở ô D7 và D9 của file Khung để tự động bốc toàn bộ Học viên dán vào trang tương ứng.")
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
                    file_name=f"DS_Lop_ChinhThuc_{file_template_in.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("⚠️ Bạn cần tải lên đủ cả file Khung Tuần và file Danh Sách Học Viên tháng để thực hiện!")

# --- TAB 4: QUẢN LÝ CƠ SỞ DỮ LIỆU CHUYÊN NGHIỆP (MỚI THÊM) ---
with tab_ql_csdl:
    st.header("🗄️ Khu vực Quản trị Cơ sở dữ liệu Nhân viên")
    
    # Hiển thị nhanh số lượng bản ghi hiện có
    conn = get_db_connection()
    current_total = conn.execute("SELECT COUNT(*) FROM nhan_vien").fetchone()[0]
    
    col_metric1, col_metric2 = st.columns(2)
    with col_metric1:
        st.metric(label="👥 Tổng số Nhân viên trong CSDL hiện tại", value=f"{current_total:,} nhân sự")
    with col_metric2:
        st.success("💡 Cấu trúc bảng này chuẩn chỉnh tương thích Schema với PostgreSQL / Supabase sau này.")
        
    st.subheader("📥 Nạp / Làm mới Danh sách Nhân viên (Master File DSNV)")
    file_dsnv = st.file_uploader("Tải lên File Danh sách nhân viên toàn công ty mới nhất (.xlsx, .xls, .csv)", type=["xlsx", "xls", "csv"])
    
    xac_nhan_ghi_de = st.checkbox("Tôi hiểu: thao tác này XÓA toàn bộ DSNV hiện có và nạp lại từ file mới (hệ thống tự sao lưu trước khi ghi đè).")
    if st.button("💾 Thực hiện Import/Cập nhật CSDL", type="primary"):
        if not file_dsnv:
            st.warning("⚠️ Vui lòng chọn file DSNV master trước khi bấm import!")
        elif not xac_nhan_ghi_de:
            st.warning("⚠️ Hãy tích ô xác nhận ghi đè trước khi import.")
        else:
            with st.spinner("Đang phân tích và đối chiếu dữ liệu cũ - mới..."):
                try:
                    df_dsnv_all = read_excel_values_only(file_dsnv)
                    kw_ma = ['mãnv', 'mnv', 'manv', 'mãsốnv', 'mãnhânviên', 'staffid']
                    kw_ten = ['họvàtên', 'họtên', 'fullname']
                    kw_dv = ['phòngban', 'phòng', 'phongban', 'phong'] 

                    parsed_new_records = []
                    seen_records = set() # Chống lặp dữ liệu trong chính file Excel

                    # 1. Đọc và lọc sạch file Excel mới
                    for s_name, df_s in df_dsnv_all.items():
                        c_ma, c_ten, c_dv = -1, -1, -1
                        for idx, row in df_s.iterrows():
                            row_raw = [str(x).strip() if x is not None else "" for x in row.values]
                            row_cleaned = [clean_header(x) for x in row_raw]
                            
                            tmp_ma = next((i for i, v in enumerate(row_cleaned) if any(k in v for k in kw_ma)), -1)
                            tmp_ten = next((i for i, v in enumerate(row_cleaned) if any(k in v for k in kw_ten)), -1)
                            tmp_dv = next((i for i, v in enumerate(row_cleaned) if any(k in v for k in kw_dv)), -1)
                            
                            if tmp_ma != -1 and tmp_ten != -1:
                                c_ma, c_ten = tmp_ma, tmp_ten
                                if tmp_dv != -1: c_dv = tmp_dv
                                continue
                                
                            if c_ma != -1 and c_ten != -1:
                                m_val = chuan_ma_nv(row_raw[c_ma])
                                t_val = row_raw[c_ten].strip()
                                dv_val_raw = row_raw[c_dv] if c_dv != -1 else ""
                                dv_chuan_hoa = chuan_hoa_don_vi(dv_val_raw)
                                
                                if m_val and m_val.lower() not in ['none', 'nan', '']:
                                    record_key = (m_val, t_val, dv_chuan_hoa)
                                    if record_key not in seen_records:
                                        seen_records.add(record_key)
                                        parsed_new_records.append({
                                            "ma_nv": m_val, "ho_ten": t_val,
                                            "don_vi_goc": dv_val_raw, "don_vi_chuan": dv_chuan_hoa
                                        })

                    if not parsed_new_records:
                        st.error("❌ Không tìm thấy dòng dữ liệu nhân sự hợp lệ nào từ file đã chọn.")
                    else:
                        file_sao_luu = tao_sao_luu_db()
                        conn_imp = get_db_connection()
                        try:
                            with conn_imp:  # một giao dịch: lỗi giữa chừng thì tự rollback, không mất dữ liệu cũ
                                # XÓA TOÀN BỘ DỮ LIỆU CŨ
                                conn_imp.execute("DELETE FROM nhan_vien")

                                # NẠP LẠI TOÀN BỘ DSNV MỚI
                                conn_imp.executemany(
                                    """
                                    INSERT INTO nhan_vien
                                    (ma_nv, ho_ten, don_vi_goc, don_vi_chuan)
                                    VALUES (?, ?, ?, ?)
                                    """,
                                    [
                                        (r["ma_nv"], r["ho_ten"], r["don_vi_goc"], r["don_vi_chuan"])
                                        for r in parsed_new_records
                                    ]
                                )
                        finally:
                            conn_imp.close()

                        tong_nv = len(parsed_new_records)
                        st.success(
                            f"✅ Đã thay thế toàn bộ dữ liệu cũ bằng {tong_nv:,} nhân sự từ file DSNV mới. "
                            f"(Bản sao lưu trước khi ghi đè: {file_sao_luu})"
                        )

                except Exception as e:
                    st.error(f"❌ Lỗi trong quá trình nạp CSDL: {str(e)}")
                                        
    st.divider()
    st.subheader("🔍 Tìm kiếm nhanh nhân sự trong CSDL nội bộ")
    search_query = st.text_input("Nhập Mã nhân viên hoặc Tên cần tra cứu thử:")
    if search_query:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT ma_nv, ho_ten, don_vi_goc, don_vi_chuan FROM nhan_vien WHERE ma_nv LIKE ? OR ho_ten LIKE ? LIMIT 20",
            (f"%{search_query}%", f"%{search_query}%")
        )
        rows = cursor.fetchall()
        if rows:
            df_preview = pd.DataFrame([dict(r) for r in rows])
            st.dataframe(df_preview, use_container_width=True)
        else:
            st.info("Không tìm thấy kết quả khớp.")
            
    conn.close()