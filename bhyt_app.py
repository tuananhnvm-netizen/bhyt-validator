import streamlit as st
import pandas as pd
from datetime import datetime
import os
import xml.etree.ElementTree as ET
import re
from openpyxl import load_workbook, Workbook
from io import BytesIO

st.set_page_config(page_title="Kiểm tra XML BHYT", layout="wide")

st.title("🔍 KIỂM TRA XML BHYT")

# ==================== TẠO TAB ====================
tab1, tab2 = st.tabs(["📋 Kiểm tra XML", "📊 So sánh Excel"])

# ==================== GLOBAL DEFINITIONS (FOR BOTH TABS) ====================

# Danh sách trường bắt buộc
MANDATORY_FIELDS_PER_FILE = {
    "Hành chính": ['MA_LK', 'MA_BN', 'HO_TEN', 'NGAY_SINH', 'GIOI_TINH', 'MA_THE_BHYT', 'MA_CSKCB', 'NGAY_VAO', 'NGAY_RA', 'MA_BENH_CHINH'],
    "Thuốc": ['MA_LK', 'MA_THUOC', 'TEN_THUOC', 'SO_LUONG', 'DON_GIA', 'THANH_TIEN_BH', 'NGAY_YL', 'NGAY_TH_YL', 'NGAY_KQ'],
    "Dịch vụ kỹ thuật": ['MA_LK', 'MA_DICH_VU', 'TEN_DICH_VU', 'SO_LUONG', 'DON_GIA_BH', 'THANH_TIEN_BH', 'NGAY_YL', 'NGAY_TH_YL', 'NGAY_KQ'],
    "Mã máy xét nghiệm": ['MA_LK', 'MA_DICH_VU', 'TEN_CHI_SO', 'NGAY_KQ'],
    "Phẫu thuật thủ thuật": ['MA_LK', 'THOI_DIEM_DBLS']
}

# ==================== TAB 1: KIỂM TRA XML ====================
with tab1:
    st.markdown("### TTTYT Tân Châu | Công cụ kiểm tra dữ liệu BHYT")

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
        except Exception as e:
            st.error(f"Lỗi khi tải ICD10.xlsx: {e}")
            return set()

    VALID_MA_BENH_CODES = load_icd10_codes()
    HANH_CHINH_DATA = {}  # Global dict for cross-file lookup

    # ====================== SIDEBAR ======================
    with st.sidebar:
        st.header("📋 Tệp ICD-10")
        uploaded_icd = st.file_uploader("Tải ICD10.xlsx", type=["xlsx"], key="icd_tab1")
        if uploaded_icd:
            with open("ICD10.xlsx", "wb") as f:
                f.write(uploaded_icd.getbuffer())
            st.success("✅ ICD10 đã tải")
            st.cache_data.clear()
            VALID_MA_BENH_CODES = load_icd10_codes()

    # ====================== UPLOAD FILES ======================
    st.subheader("📤 Chọn 5 file cần kiểm tra")
    cols = st.columns(5)
    files = [None] * 6
    file_types = ["", "Hành chính", "Thuốc", "Dịch vụ kỹ thuật", "Mã máy xét nghiệm", "Phẫu thuật thủ thuật"]

    for i in range(1, 6):
        with cols[i - 1]:
            st.markdown(f"File {file_types[i]}")
            files[i] = st.file_uploader(f"Chọn file {file_types[i]}", type=["xlsx", "xml"], key=f"f{i}_tab1")

    # ====================== HÀM KIỂM TRA ======================
    def run_all_validations_for_single_file(uploaded_file, file_type):
        all_errors = []
        file_path = f"temp_{uploaded_file.name}"
        try:
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            root = None
            if uploaded_file.name.endswith('.xlsx'):
                wb = load_workbook(file_path)
                sheet = wb.active
                headers = [str(cell.value).strip().upper() if cell.value else "" for cell in sheet[1]]

                # Check for empty headers
                if not any(h for h in headers):
                    all_errors.append({
                        'ma_lk': "Không xác định",
                        'loai_loi': "Lỗi định dạng file",
                        'mo_ta_loi': f"File '{uploaded_file.name}' không có tiêu đề cột.",
                        'truong_loi': 'Headers',
                        'nguon_file': file_type
                    })
                    return all_errors

                root = ET.Element("DATA")
                for r in range(2, sheet.max_row + 1):
                    ho_so = ET.SubElement(root, "HO_SO_KCB")
                    row_data = {}
                    for idx, h in enumerate(headers):
                        if h:
                            val = sheet.cell(r, idx + 1).value
                            # Standardize date format to YYYYMMDDHHMM
                            if isinstance(val, datetime):
                                val = val.strftime("%Y%m%d%H%M")
                            elif isinstance(val, (int, float)):
                                try:
                                    val = datetime.fromtimestamp((val - 25569) * 86400).strftime("%Y%m%d%H%M")
                                except Exception:
                                    pass
                            elem = ET.SubElement(ho_so, h)
                            elem.text = str(val).strip() if val is not None else ""
                            row_data[h] = elem.text

                    # Store Hành chính data for cross-file lookup
                    if file_type == "Hành chính":
                        ma_lk = row_data.get('MA_LK')
                        if ma_lk:
                            HANH_CHINH_DATA[ma_lk.strip().upper()] = {
                                'MA_HSBA': row_data.get('MA_HSBA', '').strip(),
                                'HO_TEN': row_data.get('HO_TEN', '').strip()
                            }

            elif uploaded_file.name.endswith('.xml'):
                tree = ET.parse(file_path)
                root = tree.getroot()

                # Store Hành chính data for cross-file lookup from XML
                if file_type == "Hành chính":
                    for ho_so in root.findall('.//HO_SO_KCB'):
                        ma_lk_tag = ho_so.find('MA_LK')
                        ma_lk = ma_lk_tag.text.strip() if ma_lk_tag is not None and ma_lk_tag.text else None
                        if ma_lk:
                            HANH_CHINH_DATA[ma_lk.strip().upper()] = {
                                'MA_HSBA': (ho_so.find('MA_HSBA').text if ho_so.find('MA_HSBA') is not None else "").strip(),
                                'HO_TEN': (ho_so.find('HO_TEN').text if ho_so.find('HO_TEN') is not None else "").strip()
                            }

            # If root is still None, something went wrong with file parsing
            if root is None:
                all_errors.append({
                    'ma_lk': "Không xác định",
                    'loai_loi': "Lỗi đọc file",
                    'mo_ta_loi': f"Không thể đọc hoặc phân tích file '{uploaded_file.name}'.",
                    'truong_loi': "File",
                    'nguon_file': file_type
                })
                return all_errors

            # === VALIDATION LOGIC START ===
            for ho_so in root.findall('.//HO_SO_KCB'):
                ma_lk_tag = ho_so.find('MA_LK')
                ma_lk = ma_lk_tag.text.strip().upper() if ma_lk_tag is not None and ma_lk_tag.text else "Không xác định"

                # Check mandatory fields
                for field in MANDATORY_FIELDS_PER_FILE.get(file_type, []):
                    tag = ho_so.find(field)
                    if tag is None or not (tag.text or "").strip():
                        all_errors.append({
                            'ma_lk': ma_lk,
                            'loai_loi': "Lỗi thiếu trường bắt buộc",
                            'mo_ta_loi': f"Trường '{field}' bị thiếu hoặc rỗng.",
                            'truong_loi': field,
                            'nguon_file': file_type
                        })

                # Validate date logic (NGAY_VAO, NGAY_RA) - specifically for Hành chính
                if file_type == "Hành chính":
                    ngay_vao_tag = ho_so.find('NGAY_VAO')
                    ngay_ra_tag = ho_so.find('NGAY_RA')

                    if ngay_vao_tag is not None and ngay_vao_tag.text and ngay_ra_tag is not None and ngay_ra_tag.text:
                        try:
                            ngay_vao = datetime.strptime(ngay_vao_tag.text.strip(), '%Y%m%d%H%M')
                            ngay_ra = datetime.strptime(ngay_ra_tag.text.strip(), '%Y%m%d%H%M')

                            if ngay_ra < ngay_vao:
                                all_errors.append({
                                    'ma_lk': ma_lk,
                                    'loai_loi': "Lỗi logic ngày",
                                    'mo_ta_loi': f"Ngày ra ({ngay_ra_tag.text}) không thể nhỏ hơn Ngày vào ({ngay_vao_tag.text}).",
                                    'truong_loi': 'NGAY_RA',
                                    'nguon_file': file_type
                                })
                        except ValueError:
                            all_errors.append({
                                'ma_lk': ma_lk,
                                'loai_loi': "Lỗi định dạng ngày",
                                'mo_ta_loi': "Ngày vào hoặc Ngày ra không đúng định dạng YYYYMMDDHHMM.",
                                'truong_loi': 'NGAY_VAO/NGAY_RA',
                                'nguon_file': file_type
                            })
                    else:
                        pass  # Covered by mandatory fields

                # Validate MA_BENH_CHINH against ICD-10
                if file_type == "Hành chính":
                    ma_benh_chinh_tag = ho_so.find('MA_BENH_CHINH')
                    if ma_benh_chinh_tag is not None and ma_benh_chinh_tag.text:
                        ma_benh_chinh = ma_benh_chinh_tag.text.strip().upper()
                        if ma_benh_chinh not in VALID_MA_BENH_CODES:
                            all_errors.append({
                                'ma_lk': ma_lk,
                                'loai_loi': "Lỗi mã bệnh ICD-10 chính",
                                'mo_ta_loi': f"Mã bệnh chính '{ma_benh_chinh}' không hợp lệ theo danh mục ICD-10.",
                                'truong_loi': 'MA_BENH_CHINH',
                                'nguon_file': file_type
                            })
                    else:
                        pass  # Covered by mandatory fields

                # Validate THOI_DIEM_DBLS format for Phẫu thuật thủ thuật
                if file_type == "Phẫu thuật thủ thuật":
                    thoi_diem_dbla_tag = ho_so.find('THOI_DIEM_DBLS')
                    if thoi_diem_dbla_tag is not None and thoi_diem_dbla_tag.text:
                        try:
                            datetime.strptime(thoi_diem_dbla_tag.text.strip(), '%Y%m%d%H%M')
                        except ValueError:
                            all_errors.append({
                                'ma_lk': ma_lk,
                                'loai_loi': "Lỗi định dạng thời điểm",
                                'mo_ta_loi': "Thời điểm bắt đầu phẫu thuật/thủ thuật không đúng định dạng YYYYMMDDHHMM.",
                                'truong_loi': 'THOI_DIEM_DBLS',
                                'nguon_file': file_type
                            })
                    else:
                        pass  # Covered by mandatory fields
            # === VALIDATION LOGIC END ===

        except Exception as e:
            all_errors.append({
                'ma_lk': "Không xác định",
                'loai_loi': f"Lỗi xử lý file '{uploaded_file.name}'",
                'mo_ta_loi': f"Lỗi chung: {str(e)}",
                'truong_loi': "File",
                'nguon_file': file_type
            })
        finally:
            if os.path.exists(file_path):
                os.remove(file_path)
        return all_errors

    # ====================== NÚT KIỂM TRA ======================
    if st.button("🚀 BẮT ĐẦU KIỂM TRA TẤT CẢ", type="primary", use_container_width=True):
        if not os.path.exists("ICD10.xlsx"):
            st.error("❌ Vui lòng tải file ICD10.xlsx ở sidebar!")
        else:
            HANH_CHINH_DATA.clear()
            all_collected_errors = []
            file_names = ["XML1 - Hành chính", "XML2 - Thuốc", "XML3 - Dịch vụ Kỹ thuật", "XML4 - Xét nghiệm", "XML5 - Phẫu thuật"]

            # Process Hành chính first to populate HANH_CHINH_DATA
            if files[1]:
                with st.spinner(f"Đang kiểm tra file {file_types[1]}..."):
                    errors = run_all_validations_for_single_file(files[1], file_types[1])
                    all_collected_errors.append((file_names[0], errors))

            # Process other files
            for i in range(2, 6):
                if files[i]:
                    with st.spinner(f"Đang kiểm tra file {file_types[i]}..."):
                        errors = run_all_validations_for_single_file(files[i], file_types[i])
                        all_collected_errors.append((file_names[i - 1], errors))

            # ====================== HIỂN THỊ KẾT QUẢ THEO TỪNG FILE ======================
            st.subheader("📊 KẾT QUẢ KIỂM TRA")

            has_error = False
            for file_name, errors in all_collected_errors:
                with st.expander(f"📋 {file_name} ({len(errors)} lỗi)", expanded=True):
                    if errors:
                        has_error = True
                        data = []
                        for err in errors:
                            ma_lk_match = re.search(r"hồ sơ .*?\|(.*?):|'(.*?)' bị thiếu", err.get('mo_ta_loi', ''))
                            ma_lk = (
                                ma_lk_match.group(1) if ma_lk_match and ma_lk_match.group(1)
                                else (ma_lk_match.group(2) if ma_lk_match and ma_lk_match.group(2)
                                      else err.get('ma_lk', "Không rõ"))
                            )

                            ma_hsba = HANH_CHINH_DATA.get(ma_lk.strip().upper(), {}).get('MA_HSBA', '')
                            ho_ten = HANH_CHINH_DATA.get(ma_lk.strip().upper(), {}).get('HO_TEN', '')

                            data.append({
                                "Mã LK": ma_lk,
                                "Mã HSBA": ma_hsba,
                                "Họ Tên": ho_ten,
                                "Loại Lỗi": err.get('loai_loi', ''),
                                "Mô Tả Lỗi Chi Tiết": err.get('mo_ta_loi', ''),
                                "Trường Lỗi": err.get('truong_loi', ''),
                                "Nguồn File": err.get('nguon_file', '')
                            })

                        df = pd.DataFrame(data)
                        st.dataframe(df, use_container_width=True)
                    else:
                        st.success("✅ Không có lỗi")

            # ====================== TẢI BÁO CÁO EXCEL ======================
            if has_error:
                full_data = []
                for file_name, errors in all_collected_errors:
                    for err in errors:
                        ma_lk_match = re.search(r"hồ sơ .*?\|(.*?):|'(.*?)' bị thiếu", err.get('mo_ta_loi', ''))
                        ma_lk = (
                            ma_lk_match.group(1) if ma_lk_match and ma_lk_match.group(1)
                            else (ma_lk_match.group(2) if ma_lk_match and ma_lk_match.group(2)
                                  else err.get('ma_lk', "Không rõ"))
                        )
                        ma_hsba = HANH_CHINH_DATA.get(ma_lk.strip().upper(), {}).get('MA_HSBA', '')
                        ho_ten = HANH_CHINH_DATA.get(ma_lk.strip().upper(), {}).get('HO_TEN', '')
                        full_data.append({
                            "Loại File": file_name,
                            "Mã LK": ma_lk,
                            "Mã HSBA": ma_hsba,
                            "Họ Tên": ho_ten,
                            "Loại Lỗi": err.get('loai_loi', ''),
                            "Mô Tả Lỗi Chi Tiết": err.get('mo_ta_loi', ''),
                            "Trường Lỗi": err.get('truong_loi', ''),
                            "Nguồn File": err.get('nguon_file', '')
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
            else:
                st.success("🎉 Không tìm thấy lỗi nào trong các file đã chọn để tạo báo cáo Excel!")
            st.caption("💡 Mỗi XML có khung kết quả riêng. Báo cáo Excel có Mã HSBA và Họ Tên.")

# ==================== TAB 2: SO SÁNH EXCEL ====================
with tab2:
    st.markdown("### 📊 So sánh dữ liệu Excel với XML / ICD-10")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("File ICD-10")
        icd_file = st.file_uploader("Upload file ICD-10.xlsx",
                                    type=["xlsx", "xls"],
                                    key="icd_tab2")

    with col2:
        st.subheader("File HSBA / Dữ liệu cần so sánh")
        hsba_file = st.file_uploader("Upload file Excel HSBA",
                                     type=["xlsx", "xls"],
                                     key="hsba_tab2")

    # Thêm nhiều file nếu cần
    st.markdown("**Các file Excel khác (tùy chọn)**")
    c1, c2, c3 = st.columns(3)
    with c1:
        ex1 = st.file_uploader("File Excel 1", type=["xlsx"], key="ex1")
    with c2:
        ex2 = st.file_uploader("File Excel 2", type=["xlsx"], key="ex2")
    with c3:
        ex3 = st.file_uploader("File Excel 3", type=["xlsx"], key="ex3")

    if st.button("🚀 BẮT ĐẦU SO SÁNH", type="primary", use_container_width=True, key="compare_tab2_button"):
        if hsba_file is None:
            st.warning("Vui lòng upload file Excel để so sánh!")
        else:
            with st.spinner("Đang đọc và so sánh dữ liệu..."):
                try:
                    df = pd.read_excel(hsba_file)
                    st.success(f"✅ Đã đọc file: **{hsba_file.name}** - {len(df)} dòng")
                    st.dataframe(df.head(15), use_container_width=True)

                    st.info("🔄 Logic so sánh sẽ được thêm sau...")
                    # Bạn có thể viết hàm so sánh ở đây sau

                except Exception as e:
                    st.error(f"Lỗi: {e}")
