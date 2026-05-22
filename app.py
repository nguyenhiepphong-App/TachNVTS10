import streamlit as st
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
import io
import re
from collections import defaultdict

st.set_page_config(page_title="Hệ Thống Xử Lý Nguyện Vọng Tuyển Sinh 10", layout="wide")

st.title("🎯 Phần Mềm Xử Lý & Thống Kê Nguyện Vọng Tuyển Sinh 10 (Bản Đỉnh Cao)")
st.markdown("""
**Hệ thống tự động xuất ra 2 loại file báo cáo chuyên nghiệp:**
1. 📄 **File Danh Sách:** Tách ngang 6 cột NV, thêm cột Tổng NV, kẻ viền đồng bộ và giữ nguyên 100% định dạng dòng tiêu đề gốc.
2. 📊 **File Thống Kê:** Tổng hợp danh sách các trường THPT/GDTX kèm số lượng đăng ký chi tiết từ NV1 đến NV6, có dòng Tổng cộng ở đáy bảng.
""")

uploaded_file = st.file_uploader("Chọn file Excel danh sách nguyện vọng gốc", type=["xlsx"])

if uploaded_file is not None:
    # --- ĐỌC VÀ XỬ LÝ FILE 1: FILE DANH SÁCH HỌC SINH ---
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active

    thin_side = Side(style='thin', color="000000")
    border_style = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    header_row_idx = None
    col_nv_idx = None
    max_col = ws.max_column
    max_row = ws.max_row

    # Tìm dòng tiêu đề bảng học sinh
    for r in range(1, min(40, max_row + 1)):
        for c in range(1, max_col + 1):
            val = str(ws.cell(row=r, column=c).value or '').lower()
            if 'nguyện vọng' in val or 'nguyenvong' in val:
                header_row_idx = r
                col_nv_idx = c
                break
        if header_row_idx:
            break

    if header_row_idx and col_nv_idx:
        start_insert_col = max_col + 1
        
        # Tạo cấu trúc đầu cột mới cho file danh sách
        new_headers = ["NV1", "NV2", "NV3", "NV4", "NV5", "NV6", "Tổng NV"]
        for i, h_text in enumerate(new_headers):
            cell = ws.cell(row=header_row_idx, column=start_insert_col + i)
            cell.value = h_text
            cell.font = Font(bold=True, name="Times New Roman", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_style

        # Bộ từ điển ngầm để phục vụ đếm dữ liệu cho FILE 2 (File Thống Kê)
        # Cấu trúc: thống_kê[tên_trường][chỉ_mục_nv_từ_0_đến_5] = số_lượng
        thong_ke_data = defaultdict(lambda: [0] * 6)
        count_processed = 0

        # Quét duyệt danh sách học sinh
        for r in range(header_row_idx + 1, max_row + 1):
            nv_text_raw = ws.cell(row=r, column=col_nv_idx).value
            nv_text = str(nv_text_raw or '').strip()
            
            lines = [line.strip() for line in nv_text.split('\n') if line.strip()]
            total_nv = len(lines)
            
            if nv_text_raw: 
                count_processed += 1

            # Tách cột dữ liệu học sinh + Thu thập số liệu thống kê trường
            for index in range(6):
                target_cell = ws.cell(row=r, column=start_insert_col + index)
                if index < total_nv:
                    line = lines[index]
                    # Bẻ chữ bỏ số thứ tự đầu dòng (ví dụ "1. THPT Liên Bảo" -> "THPT Liên Bảo")
                    match = re.match(r'^[\d\.]+\s*(.*)$', line)
                    clean_nv = match.group(1).strip() if match else line
                    
                    target_cell.value = clean_nv
                    
                    # Ghi nhận vào bộ đếm thống kê nếu dòng có chữ
                    if clean_nv:
                        thong_ke_data[clean_nv][index] += 1
                
                target_cell.font = Font(name="Times New Roman", size=11)
                target_cell.border = border_style
                target_cell.alignment = Alignment(vertical="center", wrap_text=True)

            # Điền cột Tổng NV
            total_cell = ws.cell(row=r, column=start_insert_col + 6)
            if nv_text_raw:
                total_cell.value = total_nv
            total_cell.font = Font(bold=True, name="Times New Roman", size=11)
            total_cell.border = border_style
            total_cell.alignment = Alignment(horizontal="center", vertical="center")

            # Kẻ viền cho các cột dữ liệu cũ phía trước để đồng bộ bảng
            for c in range(1, start_insert_col):
                ws.cell(row=r, column=c).border = border_style

        # Xuất dữ liệu File 1 ra bộ nhớ tạm
        output_ds = io.BytesIO()
        wb.save(output_ds)

        # --- KHỞI TẠO VÀ XỬ LÝ FILE 2: FILE THỐNG KÊ TỔNG HỢP ---
        wb_tk = openpyxl.Workbook()
        ws_tk = wb_tk.active
        ws_tk.title = "Thong Ke Nguyen Vong"
        
        # Tạo tiêu đề bảng thống kê theo chuẩn mẫu
        tk_headers = ["STT", "Tên trường", "Số NV1", "Số NV2", "Số NV3", "Số NV4", "Số NV5", "Số NV6"]
        for col_num, header_title in enumerate(tk_headers, 1):
            c_cell = ws_tk.cell(row=1, column=col_num)
            c_cell.value = header_title
            c_cell.font = Font(bold=True, name="Times New Roman", size=11)
            c_cell.alignment = Alignment(horizontal="center", vertical="center")
            c_cell.border = border_style

        # Đổ dữ liệu từ bộ đếm ngầm vào bảng thống kê
        row_tk_idx = 2
        stt_counter = 1
        
        # Sắp xếp tên trường theo thứ tự bảng chữ cái ABC cho dễ nhìn
        for truong_ten in sorted(thong_ke_data.keys()):
            if not truong_ten: continue
            
            ws_tk.cell(row=row_tk_idx, column=1, value=stt_counter).alignment = Alignment(horizontal="center")
            ws_tk.cell(row=row_tk_idx, column=2, value=truong_ten)
            
            # Điền số lượng từ NV1 đến NV6
            for nv_i in range(6):
                val_count = thong_ke_data[truong_ten][nv_i]
                # Nếu số lượng bằng 0 thì để trống ô cho thoáng bảng giống mẫu của ông
                ws_tk.cell(row=row_tk_idx, column=3 + nv_i, value=val_count if val_count > 0 else "")
                ws_tk.cell(row=row_tk_idx, column=3 + nv_i).alignment = Alignment(horizontal="center")
            
            # Kẻ border và định dạng chữ font Times New Roman
            for c_idx in range(1, 9):
                cell_obj = ws_tk.cell(row=row_tk_idx, column=c_idx)
                cell_obj.font = Font(name="Times New Roman", size=11)
                cell_obj.border = border_style
                
            stt_counter += 1
            row_tk_idx += 1

        # Tạo dòng TỔNG CỘNG ở đáy bảng thống kê
        ws_tk.cell(row=row_tk_idx, column=1, value="Tổng").font = Font(bold=True, name="Times New Roman", size=11)
        ws_tk.cell(row=row_tk_idx, column=1).alignment = Alignment(horizontal="center")
        ws_tk.cell(row=row_tk_idx, column=1).border = border_style
        ws_tk.cell(row=row_tk_idx, column=2).border = border_style # Ô tên trường hàng tổng để trống
        
        # Dùng công thức hàm SUM() của Excel để tự động cộng dồn theo cột dọc
        for nv_i in range(6):
            col_letter = openpyxl.utils.get_column_letter(3 + nv_i)
            sum_formula = f"=SUM({col_letter}2:{col_letter}{row_tk_idx-1})"
            
            sum_cell = ws_tk.cell(row=row_tk_idx, column=3 + nv_i, value=sum_formula)
            sum_cell.font = Font(bold=True, name="Times New Roman", size=11)
            sum_cell.alignment = Alignment(horizontal="center")
            sum_cell.border = border_style

        # Tự động căn chỉnh độ rộng cột Tên Trường của bảng thống kê cho vừa vặn chữ
        max_len_truong = max([len(t) for t in thong_ke_data.keys()] + [20])
        ws_tk.column_dimensions['B'].width = max_len_truong + 3
        ws_tk.column_dimensions['A'].width = 6

        # Xuất dữ liệu File 2 ra bộ nhớ tạm
        output_tk = io.BytesIO()
        wb_tk.save(output_tk)

        # --- GIAO DIỆN HIỂN THỊ NÚT TẢI FILE TRÊN STREAMLIT ---
        st.success(f"🎉 Hệ thống đã bóc tách dữ liệu và lập bảng thống kê thành công cho {count_processed} học sinh!")
        
        # Chia giao diện làm 2 cột cân xứng để đặt 2 nút tải file
        col_btn1, col_btn2 = st.columns(2)
        
        with col_btn1:
            st.info("📄 File 1: Danh sách học sinh tách cột")
            st.download_button(
                label="📥 Tải file Excel Danh Sách",
                data=output_ds.getvalue(),
                file_name="1_Danh_Sach_Tach_Cot_NV_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        with col_btn2:
            st.success("📊 File 2: Bảng tổng hợp thống kê nguyện vọng")
            st.download_button(
                label="📥 Tải file Excel Thống Kê Trường",
                data=output_tk.getvalue(),
                file_name="2_Thong_Ke_Nguyen_Vong_Truong_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Không tìm thấy cấu trúc cột 'Nguyện vọng' hợp lệ trong file gốc.")
