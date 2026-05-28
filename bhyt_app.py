import streamlit as st
import pandas as pd
from datetime import datetime
import os
import xml.etree.ElementTree as ET
import re
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="Kiểm tra XML BHYT", layout="wide")
st.title("🔍 KIỂM TRA XML BHYT")
st.markdown("**TTYT Tân Châu** | Công cụ kiểm tra dữ liệu BHYT")

# ====================== ICD-10 ======================
@st.cache_data
def load_icd10_codes():
    try:
        if os.path.exists("ICD10.xlsx"):
            wb = load_workbook("ICD10.xlsx", data_only=True)
            sheet = wb.active
            codes = set()
            for row in range(2, sheet.max_row + 1):
                val = sheet.cell(row, 2).value
                if val:
                    code = str(val).strip().upper()
                    codes.add(code)
                    if '.' in code:
                        codes.add(code.replace('.', ''))
            return codes
        return set()
    except:
        return set()

VALID_MA_BENH_CODES = load_icd10_codes()
HANH_CHINH_DATA = {}

# ====================== SIDEBAR ======================
with st.sidebar:
    st.header("📋 Tệp ICD-10")
    if st.file_uploader("Tải ICD10.xlsx", type=["xlsx"], key="icd"):
        uploaded = st.session_state.icd
        with open("ICD10.xlsx", "wb") as f:
            f.write(uploaded.getbuffer())
        st.success("✅ ICD10 đã tải")
        st.cache_data.clear()

# ====================== UPLOAD FILES ======================
st.subheader("📤 Chọn 5 file cần kiểm tra")
cols = st.columns(5)
files = [None] * 6
file_types = ["", "Hành chính", "Thuốc", "Dịch vụ kỹ thuật", "Mã máy xét nghiệm", "Phẫu thuật thủ thuật"]

for i in range(1, 6):
    with cols[i-1]:
        st.markdown(f"**XML {i}**")
        files[i] = st.file_uploader("Chọn file", type=["xlsx", "xml"], key=f"f{i}")

# ====================== HÀM KIỂM TRA ======================
def run_all_validations_for_single_file(input_file_path, file_type):
    all_errors = []
    try:
        if input_file_path.endswith('.xlsx'):
            # Đọc Excel chuyển thành XML mock
            wb = load_workbook(input_file_path)
            sheet = wb.active
            headers = [str(cell.value).strip().upper() if cell.value else "" for cell in sheet[1]]
            root = ET.Element("DATA")
            for r in range(2, sheet.max_row + 1):
                ho_so = ET.SubElement(root, "HO_SO_KCB")
                for idx, h in enumerate(headers):
                    if h:
                        val = sheet.cell(r, idx+1).value
                        elem = ET.SubElement(ho_so, h)
                        elem.text = str(val).strip() if val is not None else ""
        else:
            tree = ET.parse(input_file_path)
            root = tree.getroot()

        # Kiểm tra trường bắt buộc (rút gọn)
        for ho_so in root.findall('.//HO_SO_KCB'):
            ma_lk_tag = ho_so.find('MA_LK')
            ma_lk = ma_lk_tag.text.strip() if ma_lk_tag is not None and ma_lk_tag.text else "Không rõ"
            
            for field in MANDATORY_FIELDS_PER_FILE.get(file_type, []):
                tag = ho_so.find(field)
                if tag is None or not (tag.text or "").strip():
                    all_errors.append(f"Lỗi trường bắt buộc cho hồ sơ {file_type}|{ma_lk}: Thiếu '{field}'")

        # Lưu thông tin Hành chính
        if file_type == "Hành chính":
            for ho_so in root.findall('.//HO_SO_KCB'):
                ma_lk = ho_so.find('MA_LK').text.strip() if ho_so.find('MA_LK') is not None else None
                if ma_lk:
                    HANH_CHINH_DATA[ma_lk] = {
                        'MA_HSBA': ho_so.find('MA_HSBA').text if ho_so.find('MA_HSBA') is not None else "",
                        'HO_TEN': ho_so.find('HO_TEN').text if ho_so.find('HO_TEN') is not None else ""
                    }
    except Exception as e:
        all_errors.append(f"Lỗi xử lý file {file_type}: {str(e)}")
    
    return all_errors

# Danh sách trường bắt buộc
MANDATORY_FIELDS_PER_FILE = {
    "Hành chính": ['MA_LK', 'MA_BN', 'HO_TEN', 'NGAY_SINH', 'GIOI_TINH', 'MA_THE_BHYT', 'MA_CSKCB', 'NGAY_VAO', 'NGAY_RA', 'MA_BENH_CHINH'],
    "Thuốc": ['MA_LK', 'MA_THUOC', 'TEN_THUOC', 'SO_LUONG', 'DON_GIA', 'THANH_TIEN_BH', 'NGAY_YL', 'NGAY_TH_YL', 'NGAY_KQ'],
    "Dịch vụ kỹ thuật": ['MA_LK', 'MA_DICH_VU', 'TEN_DICH_VU', 'SO_LUONG', 'DON_GIA_BH', 'THANH_TIEN_BH', 'NGAY_YL', 'NGAY_TH_YL', 'NGAY_KQ'],
    "Mã máy xét nghiệm": ['MA_LK', 'MA_DICH_VU', 'TEN_CHI_SO', 'NGAY_KQ'],
    "Phẫu thuật thủ thuật": ['MA_LK', 'THOI_DIEM_DBLS']
}

# ====================== NÚT KIỂM TRA ======================
if st.button("🚀 BẮT ĐẦU KIỂM TRA TẤT CẢ", type="primary", use_container_width=True):
    if not os.path.exists("ICD10.xlsx"):
        st.error("❌ Vui lòng tải file ICD10.xlsx ở sidebar!")
    else:
        with st.spinner("Đang kiểm tra tất cả file..."):
            all_errors_list = []
            file_names = ["XML1 - Hành chính", "XML2 - Thuốc", "XML3 - Dịch vụ Kỹ thuật", 
                         "XML4 - Xét nghiệm", "XML5 - Phẫu thuật"]

            for i in range(1, 6):
                if files[i]:
                    temp_path = f"temp_{i}.xlsx"
                    with open(temp_path, "wb") as f:
                        f.write(files[i].getbuffer())
                    
                    errors = run_all_validations_for_single_file(temp_path, file_types[i])
                    all_errors_list.append((file_names[i-1], errors))
                    
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

            # ====================== HIỂN THỊ KẾT QUẢ THEO TỪNG FILE ======================
            st.subheader("📊 KẾT QUẢ KIỂM TRA")

            has_error = False
            for file_name, errors in all_errors_list:
                with st.expander(f"📋 {file_name} ({len(errors)} lỗi)", expanded=True):
                    if errors:
                        has_error = True
                        # Tạo DataFrame với thông tin chi tiết
                        data = []
                        for err in errors:
                            # Trích xuất MA_LK
                            ma_lk_match = re.search(r'hồ sơ .*?\|(.*?):', err)
                            ma_lk = ma_lk_match.group(1) if ma_lk_match else "Không rõ"
                            
                            ma_hsba = HANH_CHINH_DATA.get(ma_lk, {}).get('MA_HSBA', '')
                            ho_ten = HANH_CHINH_DATA.get(ma_lk, {}).get('HO_TEN', '')
                            
                            data.append({
                                "Mã LK": ma_lk,
                                "Mã HSBA": ma_hsba,
                                "Họ Tên": ho_ten,
                                "Chi Tiết Lỗi": err
                            })
                        
                        df = pd.DataFrame(data)
                        st.dataframe(df, use_container_width=True)
                    else:
                        st.success("✅ Không có lỗi")

            # ====================== TẢI BÁO CÁO EXCEL ======================
            if has_error:
                full_data = []
                for file_name, errors in all_errors_list:
                    for err in errors:
                        ma_lk_match = re.search(r'hồ sơ .*?\|(.*?):', err)
                        ma_lk = ma_lk_match.group(1) if ma_lk_match else "Không rõ"
                        ma_hsba = HANH_CHINH_DATA.get(ma_lk, {}).get('MA_HSBA', '')
                        ho_ten = HANH_CHINH_DATA.get(ma_lk, {}).get('HO_TEN', '')
                        full_data.append({
                            "Loại File": file_name,
                            "Mã LK": ma_lk,
                            "Mã HSBA": ma_hsba,
                            "Họ Tên": ho_ten,
                            "Lỗi": err
                        })
                
                df_full = pd.DataFrame(full_data)
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_full.to_excel(writer, index=False)
                output.seek(0)
                
                st.download_button(
                    "📥 Tải báo cáo Excel đầy đủ",
                    data=output,
                    file_name=f"BAO_CAO_LOI_BHYT_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.caption("💡 Mỗi XML có khung kết quả riêng. Báo cáo Excel có Mã HSBA và Họ Tên.")