# CONTEXT — App "Sổ khớp DSHV" (VIAGS TTĐT)

Tài liệu mô tả hiện trạng, lỗi cần sửa, đề xuất cải tiến và thiết kế nâng cấp nhập Kế hoạch đào tạo (KHĐT). Số dòng tham chiếu theo [app.py](app.py) tại thời điểm viết (981 dòng).

---

## 0. Trạng thái triển khai (cập nhật 07/10/2026)

**Đã làm — Tab 2 "Tạo DS lớp khung từ KHĐT"** (theo mục 6; phần còn lại của tài liệu là hiện trạng/đề xuất, Tab 4 và các lỗi P1/P2 khác để sau theo yêu cầu):
- Module mới: [khdt_core.py](khdt_core.py) (danh mục miền/loại hình, parser ngày, đọc sheet KHĐT, sắp xếp), [khdt_store.py](khdt_store.py) (bảng `khdt`, `lop_hoc`, `lop_thang`; lớp kéo dài tính vào cả tháng sau), [khdt_export.py](khdt_export.py) (điền file mẫu hoặc file nhanh), [mau_ds_lop.json](mau_ds_lop.json) (thư mục mẫu + từ khóa tên file mẫu theo loại hình).
- Luồng: upload KHĐT → tích chọn miền (TSN, DAD, NBA, VNA & đối tác; cả 4 hoặc từng miền, mỗi miền chọn sheet, mặc định theo tên sheet, NBA bỏ qua sheet "NBA rà") → phân tích, xem trước/sửa → lưu CSDL → tích chọn miền + loại hình (6 mục) → sắp xếp (theo KHĐT / thời gian tăng dần) → mỗi (miền × loại hình) một file `DS lop hoc <MIEN> <LOAI HINH> T<tháng>-<năm>.xlsx`, ZIP khi nhiều file.
- Hình thức: cột EL có dữ liệu → "Elearning + Trực tiếp", không thì "Trực tiếp"; loại hình ghi ở B8 lấy theo ô loại hình của KHĐT (thiếu thì theo danh mục).
- Kiểu xuất: (a) điền vào file mẫu (giữ logo/định dạng; ~5 giây/file, chạy song song; vượt số sheet của mẫu thì tự thêm sheet); (b) **file nhanh** chỉ có dữ liệu đúng vị trí ô như mẫu (D7/D9/D10/D11/B8 trên sheet `1..N`, Mục lục từ hàng 3), dưới 1 giây, để tự copy vào mẫu.
- Sheet `KÉO DÀI THÁNG 10` không đọc. **Đã bỏ** tính năng lọc giáo viên TTĐT MB cho miền VNA & đối tác (T9/2026: lấy đủ 166 lớp thay vì 64).
- Chuẩn hóa file xuất (chế độ điền mẫu): (1) bấm logo ở mọi sheet lớp → về `Muc luc`!A1 (openpyxl làm hỏng liên kết logo của mẫu gốc — trỏ nhầm sang file ảnh — nên `khdt_export.gan_link_logo` ghi lại trực tiếp trong gói xlsx); (2) font Mục lục thống nhất Times New Roman 12 (tiêu đề đậm, tên lớp màu liên kết); (3) vượt số sheet của mẫu (80; mẫu phục hồi 53) thì tự nhân bản sheet từ sheet `1` (kèm logo) và thêm dòng Mục lục (công thức `='n'!$D$7…` + hyperlink) — không còn tách "phần 1/2"; (4) mọi sheet lớp gộp `D7:F7`, căn trái, tăng chiều cao dòng khi tên lớp dài; (5) file Ban đầu – NVM: ô `L11` = khóa nhân viên mới (vd `VNBA26-PVHK08`) lấy từ cột "Đối tượng" của KHĐT (`khdt_core.tach_ma_khoa`; lớp không có mã khóa thì để trống). File nhanh áp dụng cùng quy tắc (trừ logo).
- Kiểm thử trên KHĐT T9/2026: TSN 209, DAD 86, NBA 238, VNA & đối tác 64 lớp (đã lọc GV); số lớp theo từng mục NBA khớp số tổng ở dòng tiêu đề, trừ vài mục chênh nhẹ do cách đếm của file gốc (C.1.2, C.2.2).
- Cũng đã sửa trong lúc làm: Tab 3 khớp lớp theo mẫu thật (vị trí bảng học viên, loại hình ở B8, so ngày theo khoảng); `chuan_ma_nv` thay `.replace('.0','')`; `.gitignore` chặn `*.db`, `backup/`.
- Đã thêm sơ bộ ở Tab 4 (ngoài phạm vi ưu tiên): sao lưu trước khi ghi đè, ô xác nhận, giao dịch.

**Cập nhật 07/10/2026 (lần 2) — tách file, tên lớp, L11, I7:O7:**
- **Tách file xuất** (`khdt_core.chia_file_xuat`): NVM của TSN/DAD/NBA → mỗi khóa (mã ở cột Đối tượng) một file, lớp không có mã gom file "KHONG MA KHOA"; miền VNA & đối tác → mỗi đối tác (mục `B.n`: AGS, NCS) một file đủ mọi loại hình, dùng mẫu `DS lOP HOC DOI TAC` (mục "Đối tác" chỉ có ở VNA & đối tác; lớp đối tác không còn nằm trong các file loại hình của VNA); loại hình khác 1 file/loại. Tên file: `DS lop hoc <MIEN> <LOAI> [<MA KHOA | DOI TAC>] T<tháng>-<năm>.xlsx`. T9/2026: 27 file cho 4 miền.
- **Tên lớp**: không ghép tên mục La Mã chỉ nơi học/hợp đồng (`Tại VCA`, `NCS - HĐĐT02`, `AGS PL28`) vào tên lớp.
- **Khóa chỉ gồm các lớp nhỏ Lý thuyết + Thực hành** (vd `3.1 Lý thuyết`, `3.2 Thực hành` dưới khóa `3`) → gộp thành **1 lớp** (tên = tên khóa, thời gian từ đầu LT đến cuối TH), như lớp định kỳ ở 3 chi nhánh; áp dụng theo cấu trúc (mọi dòng con đều là LT/TH) nên các khóa có thêm môn khác (vd `6.1.1…`) vẫn tách từng lớp. Dòng không STT nhưng có ngày (vd `Lý thuyết + kiểm tra`, `Kiểm tra lý thuyết`) được cộng vào lớp đứng trước (trước đây bị bỏ → thiếu 2 lớp phục hồi của TSN và 1 lớp của VNA).
- **L11** (NVM): `(VNBA26-PVHK08)` có ngoặc, sao chép kiểu của `K11`. **I7:O7** được gộp, căn trái (file nhanh không có nửa phải nên chỉ gộp D7:F7).
- CSDL: bảng `lop_hoc` thêm cột `doi_tac` (tự `ALTER TABLE` khi mở app); KHĐT đã lưu trước đó cần *Phân tích + Lưu lại* để tách được đối tác.
- Số lớp T9/2026 sau cập nhật: TSN 211, DAD 86, NBA 238, VNA & đối tác 164.

**Cập nhật 07/10/2026 (lần 3) — luồng một miền, lưu tự động (thay thế phần "Luồng" ở trên):**
- Bước 1: **chọn MỘT miền** (nút chọn, không tích nhiều) → upload KHĐT → chọn sheet/tháng/năm → *Phân tích và lưu KHĐT* (lưu CSDL ngay, không còn bước lưu riêng và bảng sửa; xem danh sách lớp đã đọc ở expander, chỉ đọc).
- Bước 2 chỉ hiện cho miền vừa phân tích và dùng **tháng/năm của KHĐT** (không chọn tháng xuất). TSN/DAD/NBA: tích chọn 6 mục loại hình (NVM tách mỗi khóa 1 file). **VNA & đối tác**: chọn "VNA" / "Đối tác" / "Cả hai" — phần **VNA** (lớp không thuộc đối tác) **tách file theo loại hình** như các chi nhánh (`DS lop hoc VNA <LOAI HINH>`, NVM không tách khóa; tích chọn loại hình cần xuất) để mỗi file không quá nhiều sheet; mỗi đối tác 1 file `DS lop hoc DOI TAC <tên>` (mẫu `DS lOP HOC DOI TAC`). (Đã thử gộp toàn bộ VNA vào 1 file 140 sheet rồi bỏ.)
- Chế độ sắp xếp "theo KHĐT" nay là đúng thứ tự dòng trong KHĐT (lớp kéo dài từ KHĐT tháng trước xếp cuối). T9/2026: VNA NVM 19, bổ sung 16, định kỳ 83, khác 22; AGS 17, NCS 7.
- Đã bỏ khỏi giao diện: tích chọn nhiều miền, chọn tháng xuất, mục "Đối tác" trong danh sách loại hình, bảng sửa trước khi lưu.

**Cập nhật 08/10/2026 — lớp cha không ngày + tên lớp mục "bổ sung":**
- Dòng lớp **không có ngày** mà ngay sau là các dòng con (`3` → `3.1`, `3.2`) được coi là lớp cha **kể cả khi ô Loại hình của nó có "Ban đầu"** (trước đây chỉ khi trống hoàn toàn → lớp bị bỏ và `3.1`/`3.2` tách thành 2 lớp con). Nếu các con đều là Lý thuyết/Thực hành (+ `Nhóm n`) thì gộp thành **1 lớp** mang tên lớp cha, thời gian từ đầu lý thuyết đến cuối nhóm thực hành (vd `Đóng mở cửa khoang khách tàu bay từ bên ngoài (tàu A320)`, 2–4/10/2026, 16 HV, mã khóa lấy từ dòng con). Kiểm thử trên KHĐT T10/2026 NBA (dòng 70/77/84/91).
- Lớp thuộc mục cấp 2 **"CHƯƠNG TRÌNH ĐÀO TẠO BỔ SUNG"** (`C.1.3`, `C.2.3`…) có tên chỉ là tên gốc, **không ghép** `/<tên mục La Mã>` (vd `Giáo dục dịch vụ cấp nhân viên`); các mục Chuyên môn nghiệp vụ (`x.2`) vẫn ghép `Môn học/Khóa học`.

- **Bổ sung năng định — mục chỉ gồm Lý thuyết + Thực hành** (vd `XI Điều khiển xe đầu kéo` → `1 Lý thuyết + kiểm tra`, `2 Thực hành + kiểm tra` + `Nhóm n`): gộp thành **1 lớp mang tên mục** (`Điều khiển xe đầu kéo`), thời gian từ đầu lý thuyết đến cuối thực hành, giáo viên/địa điểm/hình thức gộp, như lớp định kỳ. Điều kiện: loại hình là BSND, **tên mục bắt đầu bằng "Điều khiển xe", "Vận hành thiết bị" hoặc "Nghiệp vụ vệ sinh"** (`khdt_core.TIEN_TO_MUC_GOP`, để tránh gộp nhầm; thêm tiền tố mới tại đó), mọi lớp trong mục có tên bắt đầu "Lý thuyết…"/"Thực hành…", có cả lý thuyết lẫn thực hành (mục chỉ có một lớp "Thực hành + kiểm tra" vẫn giữ `Thực hành + kiểm tra/<mục>`; "Kiểm tra lý thuyết cuối khóa" không tính). T10/2026: NBA −3, TSN −3 (T9: TSN −4, DAD −2) lớp so với trước.

**Cần xác nhận:** file mẫu mặc định cho "Ban đầu – NVM" đang là `DS lOP HOC MAU (TỪ - T-2026)` (đoán; sửa trong mau_ds_lop.json hoặc tải mẫu ngay trên giao diện); lớp dạng `3.1` Lý thuyết / `3.2` Thực hành của đối tác đang thành 2 lớp riêng.

## 1. Tổng quan

- **Mục đích**: hỗ trợ Trung tâm đào tạo (TTĐT) lập danh sách lớp học và danh sách học viên (DSHV) từ KHĐT tháng; đối chiếu DSHV với danh sách nhân viên (DSNV).
- **Stack**: Streamlit (1 file `app.py`), SQLite (`dsnv_local.db`), pandas, openpyxl, xlrd. Python 3.12/3.13. `requirements.txt` chưa ghim phiên bản.
- **Dữ liệu**: duy nhất bảng `nhan_vien` (`id`, `ma_nv`, `ho_ten`, `don_vi_goc`, `don_vi_chuan`; index `idx_ma_nv`). Hiện ~2.799 dòng, 2.729 mã khác nhau (70 mã trùng), 197 dòng `don_vi_chuan` rỗng. **Không có bảng nào cho KHĐT, lớp học, miền, loại hình** — KHĐT chỉ sống trong `st.session_state`.
- `dsnv_cache.json` là di tích bản cũ, app không dùng.

## 2. Tính năng hiện tại (4 tab)

| Tab | Chức năng |
|---|---|
| 1. Đối chiếu DSHV | Upload file DSHV, quét từng sheet, nhận cột mã NV/họ tên/đơn vị, tra SQLite. Báo "Sai họ tên", "Sai đơn vị", "Mã NV lạ"; xuất bảng lỗi + CSV. |
| 2. Tạo khung từ KHĐT | Upload KHĐT `.xlsx` → `doc_khdt_gom_theo_tuan` (247) gom lớp theo tuần ISO → chọn tuần → `tao_file_excel_mot_tuan` (451) tạo xlsx (sheet `MucLuc` + 1 sheet/lớp đặt tên theo STT); nhiều tuần → ZIP. |
| 3. Tự động thêm học viên | Upload file khung + DS học viên tổng → `doc_dshv_ra_list` (559) + `nhoi_hoc_vien_vao_template` (607) điền học viên từ hàng 14. |
| 4. Quản lý CSDL | Nạp DSNV master (xóa toàn bảng rồi nạp lại), tìm nhanh (LIKE, 20 kết quả). |

Hàm tiện ích: `parse_date_range` (81), `week_labels_from_range` (134), `chuan_hoa_don_vi` (171), `format_khdt_display` (191), `remove_vietnamese_accents` (57).

Cách đọc KHĐT hiện nay: chọn sheet chứa "vnba"/"nba" hoặc "vna"; cột cố định theo vị trí (0 STT, 1 tên, 2 loại hình, 5 EL, 8-9 từ/đến ngày, 10 địa điểm, 11 giáo viên); sheet VNA chỉ lấy dòng có giáo viên thuộc danh sách cứng `GV_TTDT_MB` (251-255). Loại hình chuẩn hóa thành 4 chuỗi: Ban đầu, Định kỳ, Phục hồi, Bồi dưỡng kiến thức. **Chưa có TSN, DAD, đối tác; chưa tách NVM / bổ sung năng định.** Không có sắp xếp tường minh (lớp theo thứ tự dòng Excel, tuần theo thứ tự chèn).

## 3. Cấu trúc file thực tế (KHĐT T9/2026 và mẫu danh sách lớp)

### 3.1 KHĐT (mẫu V.TMM-F12a)
- Sheet: `KCQ` (khối cơ quan/thuê đối tác), `A-Đào tạo tại TSN`, `B-Đào tạo tại DAD`, `C-Đào tạo tại NBA` (+ bản `NBA rà`), `VNA & DOI TAC`, `KÉO DÀI THÁNG 10` (bỏ qua), `DSGV` (giáo viên + môn dạy), `Số lượng` (tổng hợp), `Sheet1` (toàn `#REF!`). **Miền = tên sheet.** Chưa kiểm tra chi tiết các sheet NBA, VNA & ĐỐI TÁC, KÉO DÀI.
- Cột: 0 STT · 1 Khóa · 2 Loại hình · 3 SL học viên · 4 Đối tượng/mã lớp (vd `VTSN26-CXTB03`) · 5-7 Thời lượng (EL/LT hoặc TT/TH) · 8-9 Từ/Đến ngày · 10 Địa điểm · 11 Giáo viên/Đối tác · 12 Ghi chú · 13 đơn vị (PVSĐ…) · 15-16 ngày phụ (thi lại). Hàng tiêu đề không cố định (hàng 5-6 ở TSN, 6-7 ở KCQ).
- Phân cấp (TSN): `A` khu vực → `A.1` ĐT ban đầu – nhân viên mới → `A.1.1` kiến thức chung / `A.1.2` chuyên môn nghiệp vụ / `A.1.3` chương trình đào tạo bổ sung → `I, II…` (môn/đợt) → `1, 2…` hoặc `1.2`, `-` (lớp) → `Nhóm n`. Dòng mục chứa tổng số lớp/học viên ở cột 2-3.
- Định dạng ngày: `03/9`, `S10/9`, `C10/9` (ca sáng/chiều), `Từ 12/9`, `Tháng 9`, `11-13/9`, `Trực tiếp: 7-9; 21-22/9. Trực tuyến: …`.
- Loại hình thực tế: `Ban đầu`, `Định kỳ`, `Bồi dưỡng KT`, Phục hồi; nhánh "đào tạo bổ sung" (A.1.3) cũng ghi `Ban đầu` ở cột 2 → **loại hình phải suy ra từ tiêu đề mục + cột 2**.
- Lớp đối tác/KCQ: giáo viên = `Đối tác`; lớp chia buổi (`Nội dung 4 (Buổi 1…8)`).

### 3.1b Sheet NBA (`C-Đào tạo tại NBA`, mẫu F12b) — đọc kỹ; DAD và TSN có cấu trúc giống
Chỉ dùng sheet này cho miền NBA (bỏ qua `NBA rà`). Tiêu đề ở hàng 6-7, dữ liệu từ hàng 8. Cây mục (mã mục tiền tố theo miền: TSN `A.`, DAD `B.`, NBA `C.`):

| Mục | Nội dung | Loại hình chuẩn |
|---|---|---|
| `C.1` | Đào tạo ban đầu – nhân viên mới (hàng 9) | `BAN_DAU_NVM` |
| `C.2` | Đào tạo ban đầu – bổ sung năng định (hàng 44) | `BO_SUNG_NANG_DINH` |
| `C.3` | Đào tạo định kỳ (218) | `DINH_KY` |
| `C.4` | Đào tạo phục hồi (350) | `PHUC_HOI` |
| `C.5` | Đào tạo bồi dưỡng kiến thức (377): `C.5.1` kỹ năng, `C.5.2` chuyên môn, `C.5.3` bổ sung | `BOI_DUONG_KT` |
| `C.6` | Đào tạo khác (480): `I` nhắc lại, `II` theo yêu cầu hãng HK | theo cột 2 (`Nhắc lại`, `Định kỳ`, `Ban đầu`) |

- Mỗi mục `C.1`–`C.4` có 3 mục con: `.1` kiến thức chung về HKDD, `.2` chuyên môn nghiệp vụ, `.3` chương trình đào tạo bổ sung (đây là bổ sung của chính loại hình đó, không phải "bổ sung năng định"). Bổ sung năng định chỉ là mục `C.2`.
- Dòng mục (`C.x`, `C.x.y`) có số lớp ở cột 2 và số học viên ở cột 3 → dùng đối chiếu tổng sau khi parse (vd C.2 = 65 lớp/611 HV).
- Dưới mục con: `I, II…` là môn/nghiệp vụ (có thể kèm ghi chú "Tiếp theo T8", "BS tên của lớp học", "Sửa STT", tên giáo viên rải rác cột 12-14) → lớp `1, 2…`; có dạng `-` (khóa con), `18.1`, `6.1.1` (đánh số nhiều cấp trong C.5.2), và `Nhóm n` (học viên chia nhóm thực hành, cột 3 có thể là `5hv`).
- Một lớp có thể là **dòng cha không có ngày** (vd `3 Đóng mở cửa khoang…` rồi `- Lý thuyết + kiểm tra`, `- Thực hành + kiểm tra`, `Nhóm 1/2`) → ngày/giáo viên nằm ở dòng con.
- Ô loại hình (cột 2) **không đáng tin để phân loại**: dòng `C.2.3`, `C.1.3` đều ghi `Ban đầu`, `C.3.3` ghi `Định kỳ`; dùng **mã mục cha cấp 1** (`C.1`…`C.6`) làm nguồn chính, cột 2 chỉ dùng cho `C.6`.
- Cột 4 "Đối tượng" lẫn hai nghĩa: mã khóa (`VNBA26- PVHL09`, `VNBA26-CXTB03`) hoặc nhóm đối tượng (`PVHK (QT1)`, `TKST`, `VNBA10, TTSC3, VTSN1`, `PVSĐ (DVTT)`) → lưu nguyên văn, tách mã khóa bằng regex `[A-Z]{3,4}\d{2}-\s*[A-ZĐ]+\d+`.
- Ngày: `S3/9`, `C3/9`, `04/9`, `S09/9`, `C09,10/9` (nhiều ngày cách dấu phẩy), `14/9` → `S16/9`, `28/9` → `S29/9`; ca `S/C` đứng trước ngày, ô "Đến ngày" có thể trống (lớp 1 ngày).
- Giáo viên có thể gồm nhiều người trong 1 ô (`Đỗ Thị Mỹ Bình Đỗ Thu Trang (DG+ GT)`), hoặc `GV hãng`; địa điểm `TTĐT`, `VNBA`; cột 13-14 đôi khi ghi đơn vị (`PVSĐ`) hoặc ghi chú bổ sung (`BS ngày 20/8`).
- Thông tin KHĐT cuối sheet (ghi chú ký tên) không phải lớp → dừng ở dòng trống/dòng "Phòng Tổ chức…".

### 3.2 Mẫu danh sách lớp (thư mục `…\NĂM 2026\2026 (MẪU)`, 20 file)
- Theo loại hình/miền: `DS lOP HOC MAU (2026)`, `BAN DAU BO SUNG NANG DINH`, `ĐINH KỲ`, `PHUC HOI`, `BỒI DƯỠNG KT`, `DOI TAC`, `DANH SACH VNA`; theo môn: `VNBA26-CXTB/PVHK/VSTB/ĐKVH/PVHL`, `TTSC26-BDSC`; thêm `MAU BAOCAOILOPHOC`, sơ đồ phòng học.
- Cấu trúc chung: sheet `Muc luc` (STT, Lớp học, Ngày học, Địa điểm, Giáo viên; hyperlink tới sheet lớp; logo VIAGS quay về mục lục) + mỗi lớp 1 sheet `1,2,3…` theo mẫu F10 (môn/khóa B7, loại hình/hình thức B8, thời gian B9, địa điểm B10, giáo viên B11, bảng STT/Mã NV/Họ tên/Đơn vị/Ghi chú từ hàng 12-14, nửa phải điểm danh theo ngày). Mẫu đối tác có thêm bảng kết quả F05a.
- Thư mục tháng (`2026.09 dshv`) có file theo đợt tuần (`TỪ 01-13`, `14-20`, `21-30`), theo loại hình, theo đối tác/VNA.
- **App hiện ghi layout riêng (D7, D9, B8…), không dùng chính các file mẫu này.**

## 4. Lỗi và cách khắc phục

### P0 (sai kết quả / mất dữ liệu)
1. **Tab 3 gần như không khớp lớp** — regex `Loại hình:\s*([^/]+)` (dòng 630) không khớp B8 "- Loại hình/hình thức đào tạo: …", nên `loai_hinh_kh` luôn rỗng. *Sửa*: đọc loại hình bằng regex đúng (`Loại hình/hình thức đào tạo:\s*([^/]+)`) hoặc, tốt hơn, khớp theo `lop_hoc.id` (mục 6).
2. **Năm cứng 2026** trong `format_khdt_display` (216-241), trong khi `parse_date_range` dùng `datetime.now().year`. *Sửa*: truyền năm KHĐT (đọc từ tiêu đề "THÁNG 9/2026") làm tham số duy nhất.
3. **Nạp DSNV xóa toàn bảng** không xác nhận/sao lưu; file sai = mất CSDL. *Sửa*: transaction + backup `.db` + hộp xác nhận + hiển thị so sánh trước/sau.

### P1
4. `week_labels_from_range` chỉ trả 1 tuần (lớp nhiều tuần chỉ hiện ở tuần đầu); nhãn "Tuần N" thiếu năm → trùng khi qua năm. *Sửa*: trả mọi tuần chạm, nhãn `Tuần N/năm`, sắp theo ngày ISO.
5. `.replace('.0','')` trên mã NV (737, 912) phá mã dạng "10.05"; mã có số 0 đầu bị mất. *Sửa*: đọc cột mã dưới dạng `str` (`dtype=str`), chỉ cắt đuôi `.0` bằng regex `\.0$`.
6. `parse_date_range` tách theo `-` làm hỏng `5-8-2026`; không hiểu `Tháng 9`, `Trực tiếp…; Trực tuyến…`. *Sửa*: viết lại parser trả về danh sách khoảng ngày ISO + ca; có test.
7. `chuan_hoa_don_vi`: khóa có dấu/chữ hoa là code chết (so với chuỗi đã bỏ dấu, thường); khớp chuỗi con ngắn ("hk", "kt", "cl") dễ nhầm; phụ thuộc thứ tự dict → nhiều đơn vị rơi về `.upper()` ("PHÒNG CNTT" vs "PHÒNG CÔNG NGHỆ THÔNG TIN - VIAGS KCQ" là hai giá trị). *Sửa*: bảng ánh xạ ngoài (CSV/JSON), khớp từ nguyên, ưu tiên khóa dài.
8. `remove_vietnamese_accents`: regex thủ công thiếu/trùng. *Sửa*: `unicodedata.normalize('NFD')` + xử lý `đ/Đ`.
9. `nhoi_hoc_vien_vao_template`: `break` ở lớp khớp đầu (lớp trùng tên nhận nhầm), chuỗi rỗng luôn "chứa" mọi chuỗi, so ngày bỏ qua tháng, xóa vùng cũ chỉ đến hàng 30 dù khung chỉ 15 hàng, không mở rộng định dạng khi >15 học viên.
10. Từ khóa "bsn/bsnd" trong `doc_dshv_ra_list` bị gán thành "Ban đầu" (có thể là bổ sung năng định).
11. Tab 4: `conn` dùng sau khi đã đóng khi vừa import và tìm kiếm cùng lượt chạy; kết nối không đóng khi có exception. *Sửa*: context manager.
12. Tab 2/3 không bắt lỗi (file sai → traceback); `read_excel_values_only` dùng `except:` trần. *Sửa*: try/except cụ thể + `st.error` thân thiện, kiểm tra thiếu sheet/cột.
13. Cột/hàng tiêu đề theo vị trí cố định; lọc giáo viên VNA bằng danh sách cứng 13 người thay vì sheet `DSGV` của KHĐT. *Sửa*: nhận cột theo header (tìm "Khóa đào tạo", "Từ ngày"…), đọc `DSGV`.

### P2
14. **Dữ liệu nhân sự trong git** (`dsnv_local.db`, `dsnv_cache.json`) và `.gitignore` không chặn. *Sửa*: thêm `*.db`, `*.json`, `__pycache__/`, `git rm --cached`; cân nhắc làm sạch lịch sử.
15. Không xác thực — ai mở app cũng xóa được CSDL. *Sửa*: mật khẩu/`st.secrets` cho Tab 4.
16. Formula/CSV injection: tên lớp nối vào `=HYPERLINK(...)` và ô văn bản không thoát; CSV xuất ở Tab 1. *Sửa*: tiền tố `'` cho giá trị bắt đầu `= + - @`, thoát `"`.
17. Hiệu năng Tab 1: 1 truy vấn/học viên (N+1) → nạp bảng vào dict một lần; `iterrows` → vectorize.
18. Trùng lặp: khối nhận cột Tab 1 (701-734) và Tab 4 (888-909); logic loại hình lặp (359-371, 568-571); chuẩn hóa tên file lặp (554, 831).
19. Khác: `DB_FILE` đường dẫn tương đối; hằng số cứng (font, độ rộng, `Tuần_Khác`); không log; không test; `requirements.txt` không ghim.

## 5. Cải tiến chung đề xuất

- Tách module: `db.py`, `parsers/` (khdt, dshv, ngày), `normalize.py` (đơn vị, loại hình, miền), `export.py`, `app.py` chỉ giữ UI.
- Cấu hình ngoài (`config/*.json|csv`): ánh xạ sheet→miền, từ đồng nghĩa loại hình, ánh xạ đơn vị, file mẫu theo (miền × loại hình).
- Nhận cột theo header; cảnh báo thay vì bỏ qua dòng lạ; log các dòng không nhận dạng.
- `pytest` cho parser ngày/đơn vị/loại hình với dữ liệu KHĐT T9/2026.
- Ghim phiên bản trong `requirements.txt`; thêm README ngắn.

## 6. Thiết kế nâng cấp: nhập KHĐT → danh sách lớp khung theo miền × loại hình

### 6.1 Mục tiêu
Upload KHĐT một lần → hệ thống tạo **danh sách lớp khung** cho từng **miền** (TSN, DAD, NBA, VNA & đối tác) và **loại hình** (ban đầu NVM, bổ sung năng định, định kỳ, phục hồi, bồi dưỡng kiến thức), lưu DB, lọc/sắp xếp được, xuất ra file Excel theo đúng mẫu danh sách lớp của TTĐT.

### 6.2 Mô hình dữ liệu (SQLite, thêm bảng)
- `khdt` (id, thang, nam, ten_file, ngay_nhap, ghi_chu).
- `lop_hoc` (id, khdt_id, **mien**, **loai_hinh**, stt_goc, **thu_tu_goc**, ma_khoa, ten_lop, hinh_thuc, tu_ngay, den_ngay [ISO], ca, dia_diem, giao_vien, so_hv_ke_hoach, doi_tuong, ghi_chu, sheet_goc, nhom, muc_cha).
- (Giai đoạn sau) `hoc_vien_lop` (lop_id, ma_nv, ho_ten, don_vi) để Tab 3 khớp theo `lop_hoc.id`.

### 6.3 Danh mục chuẩn
- `MIEN`: `TSN`, `DAD`, `NBA`, `VNA`, `DOI_TAC` (+ `KCQ` nếu cần). Ánh xạ sheet→miền hiển thị trên UI để người dùng sửa (NBA có 2 sheet "rà"/chính → chọn 1).
- `LOAI_HINH` (thứ tự cố định): `BAN_DAU_NVM` (NVM = nhân viên mới) → `BO_SUNG_NANG_DINH` → `DINH_KY` → `PHUC_HOI` → `BOI_DUONG_KT` → `KHAC` (mục `C.6` đào tạo khác: nhắc lại, theo yêu cầu hãng; **xuất file riêng**); mỗi loại có bảng từ đồng nghĩa (vd "ban đầu", "nhân viên mới", "bổ sung", "bsn", "định kỳ", "nhắc lại", "phục hồi", "bồi dưỡng", "BDKT"…). Loại hình suy ra từ ô cột 2 **và** tiêu đề mục cha.
- **Bổ sung năng định** nằm trong mục "Đào tạo ban đầu - bổ sung năng định" của KHĐT; vị trí mục khác nhau theo sheet (vd sheet NBA là mục `C.2`). Vì vậy không gán cứng mã mục: nhận dạng mục cấp 1 theo **nội dung tiêu đề** ("nhân viên mới", "bổ sung năng định", "định kỳ", "phục hồi", "bồi dưỡng kiến thức", "khác") và để người dùng sửa trên bảng xem trước. Mục con `x.1.3`/`x.3` "Chương trình đào tạo bổ sung" thuộc loại hình của mục cấp 1 chứa nó, **không** phải bổ sung năng định (xem 3.1b).

### 6.4 Luồng
1. Upload KHĐT → **người dùng chọn miền cần tạo** (TSN, DAD, NBA, VNA & đối tác) → hệ thống lấy dữ liệu từ **đúng sheet của miền đó** (vd chọn NBA → đọc sheet NBA). Ánh xạ miền→sheet là mặc định theo tên sheet, người dùng đổi được khi tên sheet khác (vd có sheet "NBA rà"); mỗi lần tạo chỉ dùng một sheet cho một miền, không trộn sheet.
2. Parse theo stack phân cấp `A → A.1 → A.1.x → I → 1 → Nhóm`, ghi `thu_tu_goc` theo thứ tự dòng; ngày → ISO + ca.
3. **Bảng xem trước** (có thể sửa miền/loại hình/ngày từng lớp); cảnh báo dòng không nhận dạng, ngày lạ, lớp thiếu giáo viên.
4. Lưu vào DB (ghi đè theo `khdt_id` sau xác nhận).
5. Lọc theo miền / loại hình / khoảng ngày; xem tổng hợp số lớp theo miền × loại hình (đối chiếu sheet `Số lượng`).
6. Xuất.

### 6.5 Sắp xếp
- **Theo KHĐT**: miền (TSN, DAD, NBA, VNA, Đối tác) → loại hình (thứ tự mục 6.3) → `thu_tu_goc`.
- **Thời gian tăng dần**: `tu_ngay`, `den_ngay`, rồi miền → loại hình → `thu_tu_goc`. Sắp bằng ngày ISO, không dùng nhãn tuần dạng chuỗi.
- Người dùng chọn chế độ bằng radio trên UI; áp dụng cho cả bảng xem trước và file xuất.

### 6.5b Quy tắc dựng lớp (đã chốt)
- **Nhóm gộp vào lớp cha**: các dòng `Nhóm n` (và các dòng con `Lý thuyết`/`Thực hành`) không thành lớp riêng. Lớp = 1 sheet; `tu_ngay` = ngày bắt đầu sớm nhất, `den_ngay` = ngày kết thúc muộn nhất trong các nhóm (thời gian đào tạo kéo dài từ ngày bắt đầu đến ngày kết thúc nhóm cuối). Giáo viên = tập hợp (loại trùng, giữ thứ tự) của các nhóm; số học viên = tổng các nhóm.
- **Đánh số nhiều cấp** (`6.1`, `6.1.1`…): mỗi dòng lá (`6.1.1`, `6.1.2`…) là **một lớp**; dòng `6.1` là **lớp cha** (chỉ làm tiêu đề, không tạo sheet). Tên lớp = `<tên dòng lá>/<tên lớp cha>`, vd `Thực hành + kiểm tra/Nghiệp vụ cán bộ kiểm soát tại quầy`. Cùng quy tắc với tên ghép của VNA hiện có (`Lý thuyết + kiểm tra/…`). Dòng lá không có loại hình/ngày/số lượng (vd `6.1.3 Kỹ năng chăm sóc khách hàng…` chỉ có ghi chú) → không tạo lớp, đưa vào cảnh báo "dòng bỏ qua" ở bảng xem trước.
- Lớp có các dòng con `-` (vd `3 Đóng mở cửa khoang…` → `- Lý thuyết + kiểm tra`, `- Thực hành + kiểm tra`, kèm `Nhóm n`): **gộp cả các dòng `-` và `Nhóm` vào 1 lớp = 1 sheet**, giống các lớp định kỳ có lý thuyết + thực hành. Tên lớp = tên dòng cha (`Đóng mở cửa khoang…`), không ghép tên dòng con. Thời gian = từ ngày bắt đầu sớm nhất (lý thuyết) đến ngày kết thúc muộn nhất (nhóm thực hành cuối); giáo viên/địa điểm gộp như trên.
- Phân biệt hai trường hợp: dòng `-` có **dòng cha có tên riêng ngay trên nó** (không số liệu) → gộp vào lớp cha như trên; mã `6.1.1` (đánh số nhiều cấp trong `C.5.2`) → **mỗi dòng lá là lớp riêng, 1 sheet riêng**, có ngày, giáo viên, địa điểm bình thường, chỉ khác tên lớp ghép `<dòng lá>/<lớp cha 6.1>` (vd `Thực hành + kiểm tra/Nghiệp vụ cán bộ kiểm soát tại quầy`), cùng kiểu tên với các lớp ban đầu, bổ sung năng định.
- Đặt `thu_tu_goc` theo dòng lớp (không theo dòng nhóm) để sắp xếp "theo KHĐT" ổn định.

### 6.6 Xuất
- **Mỗi loại hình một file danh sách lớp khung riêng** (6 file cho mỗi miền: NVM, bổ sung năng định, định kỳ, phục hồi, bồi dưỡng KT, khác/`C.6`). Loại hình không có lớp trong tháng thì không tạo file (hiện cảnh báo). Tên file gợi ý: `DS lop hoc <MIEN> <LOAI HINH> T<thang>-<nam>.xlsx`; khi chọn nhiều loại hình → ZIP.
- Chọn file mẫu theo (miền × loại hình) từ `2026 (MẪU)` (cấu hình trong `config`), điền sheet `Muc luc` và nhân bản sheet lớp từ mẫu (giữ logo, định dạng, công thức).
- Chế độ xuất: 1 file / (miền, loại hình); hoặc gom theo tuần (`Tuần N/năm`; lớp nhiều tuần xuất hiện ở mọi tuần chạm); ZIP khi nhiều file.
- Dùng lại: logic Mục lục/hyperlink của `tao_file_excel_mot_tuan` (451), `norm_name` và so khớp tên ở `nhoi_hoc_vien_vao_template`.

### 6.7 Điểm bám trong code
- Bảng `lop_hoc` thêm `thang_hien_thi` hoặc bảng nối `lop_thang` (lop_id, thang, nam) để một lớp kéo dài xuất hiện ở nhiều tháng.
- Thay `dict_theo_tuan` (đầu ra `doc_khdt_gom_theo_tuan`) bằng danh sách phẳng các lớp; mở rộng `lop_info` (~440) với miền, loại hình chuẩn, `thu_tu_goc`, ngày ISO.
- Tab 2 đọc từ DB thay vì `session_state`.

### 6.8 Lộ trình
0. Sửa lỗi P0 (mục 4).
1. Schema + danh mục + parser mới (kèm test với KHĐT T9/2026).
2. UI xem trước / lọc / sắp xếp.
3. Xuất theo file mẫu.
4. Nối Tab 3 theo `lop_hoc.id`; xử lý P1/P2 còn lại.

## 7. Câu hỏi mở (cần người dùng xác nhận)
- ~~NBA/DAD có cấu trúc giống TSN~~ → đã chốt: giống; chỉ dùng sheet NBA cho miền NBA (xem 3.1b).
- ~~Lớp `Nhóm n`~~ → đã chốt: gộp vào lớp cha, thời gian từ ngày bắt đầu đến ngày kết thúc nhóm cuối (6.5b).
- ~~Đánh số `6.1.1`~~ → đã chốt: mỗi dòng lá là 1 lớp, `6.1` là lớp cha, tên `<dòng lá>/<lớp cha>` (6.5b).
- ~~Mục `C.6`~~ → đã chốt: file riêng; mỗi loại hình một file (6.6).
- ~~Dòng `-` lý thuyết/thực hành~~ → đã chốt: gộp vào 1 lớp, 1 sheet (6.5b).
- ~~File mẫu cho loại hình `KHAC`~~ → đã chốt: dùng `DS lOP HOC MAU (2026)`.
- ~~`KÉO DÀI THÁNG 10`~~ → đã chốt: **bỏ qua, không đọc** sheet này.
- Một lớp có nhiều `Nhóm` → xuất thành 1 sheet hay mỗi nhóm 1 sheet?
- ~~Sheet NBA hay bản rà~~ → đã chốt: người dùng chọn miền, hệ thống dùng sheet tương ứng (mục 6.4). Còn cần chốt: khi có 2 sheet cùng miền (`NBA` và `NBA rà`) thì mặc định chọn sheet nào, hay luôn hỏi.
- ~~Lớp kéo dài sang tháng sau~~ → đã chốt: tính vào **tháng bắt đầu và cả tháng sau** (lớp xuất hiện ở danh sách của cả hai tháng; lớp trải qua nhiều tháng thì xuất hiện ở mọi tháng chạm). Cần lưu cờ `keo_dai` và tháng gốc để không đếm trùng khi thống kê tổng số lớp. Sheet `KÉO DÀI THÁNG 10` **không đọc**; việc lớp xuất hiện ở tháng sau được tính từ ngày bắt đầu/kết thúc của lớp trong sheet miền.
- ~~NVM~~ → đã chốt: nhân viên mới. ~~Vị trí bổ sung năng định~~ → đã chốt: mục "Đào tạo ban đầu - bổ sung năng định" (vd `C.2` ở NBA), nhận dạng theo tiêu đề mục.
- File mẫu nào ứng với từng (miền × loại hình); đối tác/VNA có dùng riêng mẫu `DOI TAC`/`DANH SACH VNA` không?
