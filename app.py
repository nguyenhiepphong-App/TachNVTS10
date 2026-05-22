import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
import io
import re

st.set_page_config(page_title="Tách Cột Nguyện Vọng Tuyển Sinh", layout="wide")

st.title("🎯 Phần mềm Tách Cột Nguyện Vọng Tuyển Sinh 10 (Giữ Nguyên Định Dạng Gốc)")
st.markdown("""
**Tính năng nâng cao:**
* **Giữ nguyên vẹn 100% file gốc:** Tiêu đề, màu sắc, font chữ, các dòng thông tin phía trên được bảo toàn hoàn hảo.
* Hệ thống chỉ chèn thêm các cột **NV1, NV2, NV3, NV4, NV5, NV6** vào ngay sau cột cuối cùng của bảng danh sách học sinh.
""")

uploaded_file = st.file_uploader("Chọn file Excel danh sách nguyện vọng gốc", type=["xlsx"])

if uploaded_file is not None:
    # Đọc trực tiếp bằng openpyxl để giữ nguyên định dạng, font chữ, màu sắc
    wb = openpyxl.load_workbook(uploaded_file)
    ws = wb.active

    # Bước 1: Tìm dòng tiêu đề bảng (chứa chữ Họ tên / Nguyện vọng) và cột Nguyện vọng
    header_row_idx = None
    col_nv_idx = None
    max_col = ws.max_column
    max_row = ws.max_row

    # Quét để tìm vị trí bảng dữ liệu học sinh
    for r in range(1, min(30, max_row + 1)):
        for c in range(1, max_col + 1):
            val = str(ws.cell(row=r, column=c).value or '').lower()
            if 'nguyện vọng' in val or 'nguyenvong' in val:
                header_row_idx = r
                col_nv_idx = c
                break
        if header_row_idx:
            break

    if header_row_idx and col_nv_idx:
        # Xác định vị trí chèn thêm 6 cột nguyện vọng (ngay sau cột cuối cùng hiện tại của file)
        start_insert_col = max_col + 1
        
        # Điền tiêu đề cho 6 cột mới ở dòng tiêu đề bảng học sinh
        for i in range(6):
            cell = ws.cell(row=header_row_idx, column=start_insert_col + i)
            cell.value = f"NV{i+1}"
            cell.font = openpyxl.styles.Font(bold=True, name="Times New Roman", size=11)
            # Sao chép định dạng đường viền (border) từ ô bên cạnh nếu có
            ref_cell = ws.cell(row=header_row_idx, column=max_col)
            if ref_cell.border:
                cell.border = openpyxl.styles.Border(
                    left=ref_cell.border.left,
                    right=ref_cell.border.right,
                    top=ref_cell.border.top,
                    bottom=ref_cell.border.bottom
                )

        # Bước 2: Duyệt từng dòng học sinh từ sau dòng tiêu đề đến hết file
        count_processed = 0
        for r in range(header_row_idx + 1, max_row + 1):
            nv_text = str(ws.cell(row=r, column=col_nv_idx).value or '').strip()
            
            # Nếu dòng này trống trơn (không có dữ liệu học sinh), bỏ qua để giữ nguyên cấu trúc dòng trống
            if not nv_text:
                continue
                
            count_processed += 1
            lines = [line.strip() for line in nv_text.split('\n') if line.strip()]
            
            # Tách và điền vào 6 cột mới
            for index, line in enumerate(lines[:6]):
                match = re.match(r'^[\d\.]+\s*(.*)$', line)
                if match:
                    clean_nv = match.group(1).strip()
                else:
                    clean_nv = line
                
                target_cell = ws.cell(row=r, column=start_insert_col + index)
                target_cell.value = clean_nv
                # Giữ font giống các ô khác trong dòng
                target_cell.font = openpyxl.styles.Font(name="Times New Roman", size=11)
                
                # Áp border cho đẹp đội hình
                ref_cell = ws.cell(row=r, column=max_col)
                if ref_cell.border:
                    target_cell.border = openpyxl.styles.Border(
                        left=ref_cell.border.left,
                        right=ref_cell.border.right,
                        top=ref_cell.border.top,
                        bottom=ref_cell.border.bottom
                    )

        st.success(f"🎉 Đã xử lý xong! Giữ nguyên định dạng gốc và thêm 6 cột NV thành công cho {count_processed} dòng dữ liệu.")
        
        # Xuất file trả về cho người dùng
        output = io.BytesIO()
        wb.save(output)
        
        st.download_button(
            label="📥 Tải file Excel kết quả (Giữ nguyên định dạng gốc)",
            data=output.getvalue(),
            file_name="danh_sach_giu_nguyen_dinh_dang_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Lỗi kịch bản: Không tìm thấy cột 'Nguyện vọng' trong bảng dữ liệu để tiến hành bóc tách.")
