import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Tách Cột Nguyện Vọng Tuyển Sinh", layout="wide")

st.title("🎯 Phần mềm Tách Cột Nguyện Vọng Tuyển Sinh 10")
st.markdown("""
**Hướng dẫn sử dụng:**
1. Tải lên file Excel danh sách đăng ký dự thi gốc.
2. Hệ thống tự động bóc tách các nguyện vọng trong ô gộp và điền ngang vào các cột **NV1, NV2, NV3, NV4, NV5, NV6**.
3. **Thứ tự học sinh và cấu trúc dòng trong danh sách được giữ nguyên vẹn 100%.**
""")

uploaded_file = st.file_uploader("Chọn file Excel danh sách nguyện vọng gốc", type=["xlsx"])

if uploaded_file is not None:
    # Đã sửa: Ép buộc dùng engine openpyxl tại đây để trị dứt điểm lỗi ImportError
    raw_df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
    
    header_row_idx = 0
    for idx, row in raw_df.iterrows():
        if row.astype(str).str.contains('Nguyện vọng|Họ và tên', case=False).any():
            header_row_idx = idx
            break
            
    df = pd.read_excel(uploaded_file, skiprows=header_row_idx, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    
    col_hoten = next((c for c in df.columns if 'họ và tên' in c.lower() or 'hoten' in c.lower()), None)
    col_nv = next((c for c in df.columns if 'nguyện vọng' in c.lower() or 'nguyenvong' in c.lower()), None)
    
    if col_hoten and col_nv:
        df = df[df[col_hoten].notna()]
        df = df[df[col_hoten].astype(str).str.strip() != '']
        
        for i in range(1, 7):
            df[f'NV{i}'] = ""
            
        for idx, row in df.iterrows():
            nv_text = str(row[col_nv])
            lines = [line.strip() for line in nv_text.split('\n') if line.strip()]
            
            for index, line in enumerate(lines[:6]):
                match = re.match(r'^[\d\.]+\s*(.*)$', line)
                if match:
                    clean_nv = match.group(1).strip()
                else:
                    clean_nv = line
                    
                df.at[idx, f'NV{index+1}'] = clean_nv

        st.success(f"🎉 Xử lý thành công! Đã tách dữ liệu thành 6 cột nguyện vọng ngang cho {len(df)} thí sinh.")
        st.write("Xem trước bảng kết quả (Thứ tự danh sách giữ nguyên):")
        st.dataframe(df.head(20))
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
        st.download_button(
            label="📥 Tải file Excel kết quả (Có cột NV1...NV6)",
            data=output.getvalue(),
            file_name="danh_sach_tach_ngang_nguyen_vong_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Lỗi: Không tìm thấy cột 'Họ và tên' hoặc cột 'Nguyện vọng' trong file Excel của bạn.")
