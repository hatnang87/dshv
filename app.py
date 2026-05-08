import streamlit as st
import pandas as pd
import openpyxl
import io

# Cấu hình trang
st.set_page_config(page_title="🔍 Công cụ Đối chiếu Danh sách VIAGS", layout="wide")

# CSS tùy chỉnh để giao diện sạch sẽ, chuyên nghiệp
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    .status-box { padding: 20px; border-radius: 10px; border: 1px solid #dee2e6; background-color: white; }
    </style>
""", unsafe_allow_html=True)

def clean_header(val):
    """Chuẩn hóa tiêu đề để tìm kiếm chính xác không phân biệt dấu/khoảng trắng"""
    return "".join(str(val).lower().split())

def read_excel_values_only(uploaded_file):
    """Hỗ trợ đọc đa định dạng: .xlsx, .xls, .csv và lấy dữ liệu thô"""
    name = uploaded_file.name.lower()
    if name.endswith('.csv'):
        try: return {"CSV": pd.read_csv(uploaded_file, header=None, encoding='utf-8')}
        except:
            uploaded_file.seek(0)
            return {"CSV": pd.read_csv(uploaded_file, header=None, encoding='utf-8-sig')}
    elif name.endswith('.xls'):
        # Cần cài thư viện xlrd để đọc .xls
        return pd.read_excel(uploaded_file, sheet_name=None, header=None, engine='xlrd')
    else:
        # Dùng openpyxl để lấy giá trị chính xác (data_only=True) tránh lỗi công thức
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        res = {}
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            data = [list(row) for row in ws.iter_rows(values_only=True)]
            res[sheet_name] = pd.DataFrame(data)
        return res

# Giao diện chính
st.title("🔍 Radar Đối chiếu Danh sách Học viên")
st.info("💡 Hướng dẫn: Radar sẽ tự động quét tiêu đề (Mã NV, Họ tên) dù ở bất kỳ vị trí nào trong file.")

with st.container():
    col_file1, col_file2 = st.columns(2)
    with col_file1:
        file_dsnv = st.file_uploader("1. File Danh sách nhân viên gốc (DSNV.xlsx/xls)", type=["xlsx", "xls", "csv"])
    with col_file2:
        file_dshv = st.file_uploader("2. File Danh sách học viên cần kiểm tra", type=["xlsx", "xls", "csv"])

if st.button("🚀 Bắt đầu quét đối chiếu", type="primary"):
    if not file_dsnv or not file_dshv:
        st.warning("⚠️ Vui lòng tải lên đầy đủ cả 2 file!")
    else:
        with st.spinner("Đang quét dữ liệu..."):
            try:
                # 1. Xử lý File Gốc (Xây dựng từ điển Mã -> Tên chuẩn)
                df_dsnv_all = read_excel_values_only(file_dsnv)
                dict_nv_chuan = {}
                kw_ma = ['mãnv', 'mnv', 'manv', 'mãsốnv', 'mãnhânviên', 'staffid']
                kw_ten = ['họvàtên', 'họtên', 'fullname']

                for s_name, df_s in df_dsnv_all.items():
                    c_ma, c_ten = -1, -1
                    for idx, row in df_s.iterrows():
                        row_raw = [str(x).strip() if x is not None else "" for x in row.values]
                        row_cleaned = [clean_header(x) for x in row_raw]
                        
                        # Tìm tọa độ cột tiêu đề
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

                # 2. Xử lý File Học viên (Đối chiếu)
                results = []
                total_checked = 0
                dict_hv_sheets = read_excel_values_only(file_dshv)
                
                for sheet_name, df_sheet in dict_hv_sheets.items():
                    c_ma, c_ten = -1, -1
                    current_class = "N/A"
                    for idx, row in df_sheet.iterrows():
                        row_raw = [str(x).strip() if x is not None else "" for x in row.values]
                        row_cleaned = [clean_header(x) for x in row_raw]
                        
                        # --- A. NHẬN DIỆN THÔNG TIN LỚP (ƯU TIÊN CỘT A + G, H) ---
                        cell_0 = row_raw[0] # Cột A
                        cell_6 = row_raw[6] if len(row_raw) > 6 else "" # Cột G (Bắt đầu)
                        cell_7 = row_raw[7] if len(row_raw) > 7 else "" # Cột H (Kết thúc)
                        
                        # 1. TRƯỜNG HỢP LỚP THÔNG THƯỜNG: Tên lớp (A) và Ngày (G, H) cùng 1 dòng
                        # Điều kiện: Cột A có chữ dài và Cột G có định dạng ngày (chứa dấu /)
                        if cell_0 and "/" in str(cell_6):
                            current_class = f"{cell_0.split(chr(10))[0]} [{cell_6} - {cell_7}]"
                            pending_class_name = "" # Reset biến tạm vì đã chốt được lớp
                            
                        # 2. TRƯỜNG HỢP LỚP PHỤC HỒI: Tên lớp (A) nằm trên, Ngày nằm dưới
                        elif cell_0 and len(cell_0) > 10 and not any(k in clean_header(cell_0) for k in kw_ma + kw_ten):
                            # Lưu tạm tên lớp vào hàng chờ
                            pending_class_name = cell_0.split(chr(10))[0]
                            
                        # 3. NHẶT NGÀY CHO LỚP PHỤC HỒI (Dòng Lý thuyết/Thực hành)
                        if ("lý thuyết" in cell_0.lower() or "thực hành" in cell_0.lower()) and "/" in str(cell_6):
                            if pending_class_name:
                                # Lấy ngày ở dòng phụ này đắp vào tên lớp đang chờ
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