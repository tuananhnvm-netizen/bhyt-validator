import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re


# ====================== HELPER FUNCTIONS ======================

def read_excel_file(uploaded_file, sheet_name=0):
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, dtype=str)
        df = df.fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        return df, None
    except Exception as e:
        return None, str(e)


def get_sheet_names(uploaded_file):
    try:
        uploaded_file.seek(0)
        xl = pd.ExcelFile(uploaded_file)
        return xl.sheet_names
    except Exception:
        return ["Sheet1"]


def to_excel_bytes(df, sheet_name="Sheet1", highlight_cols=None, highlight_rows=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]
        # Style header
        header_fill = PatternFill("solid", fgColor="1F4E79")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        # Highlight rows
        if highlight_rows:
            yellow = PatternFill("solid", fgColor="FFFF00")
            red_fill = PatternFill("solid", fgColor="FFD7D7")
            for row_idx in highlight_rows:
                excel_row = row_idx + 2
                for col in range(1, len(df.columns) + 1):
                    ws.cell(row=excel_row, column=col).fill = red_fill
        # Auto column width
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)
    output.seek(0)
    return output


def multi_sheet_excel(sheets_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sname, df in sheets_dict.items():
            df.to_excel(writer, index=False, sheet_name=sname[:31])
            wb = writer.book
            ws = writer.sheets[sname[:31]]
            header_fill = PatternFill("solid", fgColor="1F4E79")
            header_font = Font(bold=True, color="FFFFFF")
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            for col in ws.columns:
                max_len = max((len(str(cell.value or "")) for cell in col), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)
    output.seek(0)
    return output


# ====================== FEATURE FUNCTIONS ======================

def feature_compare_two_files(df1, df2, key_col1, key_col2, compare_cols1, compare_cols2):
    """So sánh chi tiết 2 file theo cột khóa"""
    results = {}

    df1_keys = set(df1[key_col1].str.strip().str.upper())
    df2_keys = set(df2[key_col2].str.strip().str.upper())

    only_in_f1 = df1_keys - df2_keys
    only_in_f2 = df2_keys - df1_keys
    in_both = df1_keys & df2_keys

    results["only_in_file1"] = df1[df1[key_col1].str.strip().str.upper().isin(only_in_f1)].copy()
    results["only_in_file2"] = df2[df2[key_col2].str.strip().str.upper().isin(only_in_f2)].copy()

    # So sánh dòng khớp key
    diff_rows = []
    if compare_cols1 and compare_cols2 and len(compare_cols1) == len(compare_cols2):
        df1_idx = df1.set_index(df1[key_col1].str.strip().str.upper())
        df2_idx = df2.set_index(df2[key_col2].str.strip().str.upper())
        for key in in_both:
            row1 = df1_idx.loc[key] if key in df1_idx.index else None
            row2 = df2_idx.loc[key] if key in df2_idx.index else None
            if row1 is None or row2 is None:
                continue
            diffs = []
            for c1, c2 in zip(compare_cols1, compare_cols2):
                v1 = str(row1[c1]).strip() if c1 in row1.index else ""
                v2 = str(row2[c2]).strip() if c2 in row2.index else ""
                if v1 != v2:
                    diffs.append(f"{c1}: [{v1}] ≠ [{v2}]")
            if diffs:
                diff_rows.append({
                    "Khóa": key,
                    "Điểm khác biệt": " | ".join(diffs),
                    "Số trường khác": len(diffs)
                })
    results["differences"] = pd.DataFrame(diff_rows) if diff_rows else pd.DataFrame()
    results["stats"] = {
        "Tổng file 1": len(df1),
        "Tổng file 2": len(df2),
        "Chỉ có trong file 1": len(only_in_f1),
        "Chỉ có trong file 2": len(only_in_f2),
        "Có trong cả 2": len(in_both),
        "Dòng có sự khác biệt": len(diff_rows)
    }
    return results


def feature_find_duplicates(df, subset_cols, keep="first"):
    """Tìm và xử lý dữ liệu trùng"""
    dup_mask = df.duplicated(subset=subset_cols, keep=False)
    duplicates = df[dup_mask].copy()
    duplicates["__Nhóm_trùng__"] = duplicates.groupby(subset_cols).ngroup() + 1
    duplicates = duplicates.sort_values("__Nhóm_trùng__")

    cleaned = df.drop_duplicates(subset=subset_cols, keep=keep).copy()
    removed = df[df.duplicated(subset=subset_cols, keep=keep)].copy()
    return duplicates, cleaned, removed


def feature_vlookup(df_main, df_lookup, main_key, lookup_key, lookup_cols):
    """VLOOKUP: ghép dữ liệu từ file tra cứu vào file chính"""
    df_lookup_sub = df_lookup[[lookup_key] + lookup_cols].drop_duplicates(subset=[lookup_key])
    merge_map = {c: f"{c}_lookup" for c in lookup_cols if c in df_main.columns}
    df_lookup_sub = df_lookup_sub.rename(columns=merge_map)
    result = df_main.merge(
        df_lookup_sub,
        left_on=main_key,
        right_on=lookup_key,
        how="left"
    )
    not_found = result[result[lookup_key].isna()][main_key].tolist()
    return result, not_found


def feature_filter_advanced(df, conditions):
    """
    Lọc nâng cao với nhiều điều kiện
    conditions: list of dict {col, operator, value}
    """
    mask = pd.Series([True] * len(df), index=df.index)
    for cond in conditions:
        col = cond.get("col")
        op = cond.get("operator")
        val = cond.get("value", "")
        if col not in df.columns:
            continue
        series = df[col].astype(str).str.strip()
        if op == "=":
            mask &= series.str.upper() == str(val).strip().upper()
        elif op == "≠":
            mask &= series.str.upper() != str(val).strip().upper()
        elif op == "chứa":
            mask &= series.str.contains(str(val), case=False, na=False)
        elif op == "không chứa":
            mask &= ~series.str.contains(str(val), case=False, na=False)
        elif op == "bắt đầu bằng":
            mask &= series.str.startswith(str(val), na=False)
        elif op == "trống":
            mask &= (series == "") | (series.str.upper() == "NAN")
        elif op == "không trống":
            mask &= (series != "") & (series.str.upper() != "NAN")
        elif op == ">":
            try:
                mask &= pd.to_numeric(df[col], errors="coerce") > float(val)
            except Exception:
                pass
        elif op == "<":
            try:
                mask &= pd.to_numeric(df[col], errors="coerce") < float(val)
            except Exception:
                pass
    return df[mask].copy()


def feature_summary_stats(df, group_col, agg_cols):
    """Thống kê nhóm: đếm, tổng, trung bình"""
    agg_dict = {}
    for col in agg_cols:
        df[col + "_num"] = pd.to_numeric(df[col], errors="coerce")
        agg_dict[col + "_num"] = ["count", "sum", "mean", "min", "max"]
    grouped = df.groupby(group_col).agg(agg_dict).round(2)
    grouped.columns = ["_".join(c).strip() for c in grouped.columns]
    return grouped.reset_index()


def feature_check_missing_values(df):
    """Kiểm tra giá trị thiếu/trống"""
    report = []
    for col in df.columns:
        empty = ((df[col].astype(str).str.strip() == "") | (df[col].astype(str).str.upper() == "NAN")).sum()
        if empty > 0:
            report.append({
                "Cột": col,
                "Số dòng trống": empty,
                "Tỷ lệ (%)": round(empty / len(df) * 100, 1)
            })
    return pd.DataFrame(report)


def feature_normalize_data(df, cols_to_normalize):
    """Chuẩn hóa dữ liệu: trim, upper/lower, xóa ký tự đặc biệt"""
    result = df.copy()
    for col in cols_to_normalize:
        if col in result.columns:
            result[col] = result[col].astype(str).str.strip()
    return result


def feature_cross_check(df1, df2, check_col1, check_col2):
    """Kiểm tra chéo: tìm giá trị của cột A (file1) không có trong cột B (file2)"""
    vals1 = set(df1[check_col1].astype(str).str.strip().str.upper())
    vals2 = set(df2[check_col2].astype(str).str.strip().str.upper())
    missing_in_f2 = vals1 - vals2
    missing_in_f1 = vals2 - vals1
    rows_missing = df1[df1[check_col1].astype(str).str.strip().str.upper().isin(missing_in_f2)].copy()
    rows_missing["__Ghi_chú__"] = "Không có trong File 2"
    return rows_missing, len(missing_in_f2), len(missing_in_f1)


def feature_concat_files(dfs, add_source=True, source_names=None):
    """Gộp nhiều file thành 1"""
    all_dfs = []
    for i, df in enumerate(dfs):
        d = df.copy()
        if add_source:
            name = source_names[i] if source_names and i < len(source_names) else f"File {i+1}"
            d.insert(0, "__Nguồn__", name)
        all_dfs.append(d)
    return pd.concat(all_dfs, ignore_index=True, sort=False)


# ====================== MAIN TAB FUNCTION ======================

def render_excel_compare_tab():
    st.markdown("### 📊 Công cụ xử lý Excel nâng cao")

    # Upload section
    st.markdown("#### 📁 Tải lên file Excel")
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        file1 = st.file_uploader("📂 File 1 (File chính / File nguồn)", type=["xlsx", "xls"], key="ec_file1")
    with col_u2:
        file2 = st.file_uploader("📂 File 2 (File tra cứu / File so sánh)", type=["xlsx", "xls"], key="ec_file2")

    col_u3, col_u4, col_u5 = st.columns(3)
    with col_u3:
        file3 = st.file_uploader("📂 File 3 (tùy chọn)", type=["xlsx"], key="ec_file3")
    with col_u4:
        file4 = st.file_uploader("📂 File 4 (tùy chọn)", type=["xlsx"], key="ec_file4")
    with col_u5:
        file5 = st.file_uploader("📂 File 5 (tùy chọn)", type=["xlsx"], key="ec_file5")

    # Load dataframes
    dfs = {}
    file_map = {"File 1": file1, "File 2": file2, "File 3": file3, "File 4": file4, "File 5": file5}
    for fname, fobj in file_map.items():
        if fobj:
            fobj.seek(0)
            sheets = get_sheet_names(fobj)
            fobj.seek(0)
            if len(sheets) > 1:
                chosen = st.selectbox(f"Chọn sheet của {fname}:", sheets, key=f"sheet_{fname}")
            else:
                chosen = sheets[0]
            fobj.seek(0)
            df, err = read_excel_file(fobj, sheet_name=chosen)
            if err:
                st.error(f"❌ Lỗi đọc {fname}: {err}")
            else:
                dfs[fname] = df
                st.success(f"✅ {fname}: **{fobj.name}** — {len(df):,} dòng × {len(df.columns)} cột")

    if not dfs:
        st.info("👆 Vui lòng tải lên ít nhất 1 file Excel để bắt đầu.")
        return

    st.divider()

    # Feature selector
    st.markdown("#### ⚙️ Chọn tác vụ cần thực hiện")
    feature = st.selectbox("🔧 Tác vụ:", [
        "1️⃣  So sánh 2 file — Tìm điểm khác biệt",
        "2️⃣  Tìm & lọc dữ liệu TRÙNG",
        "3️⃣  VLOOKUP — Ghép dữ liệu từ file tra cứu",
        "4️⃣  Lọc dữ liệu nâng cao (nhiều điều kiện)",
        "5️⃣  Kiểm tra chéo 2 cột (Cross-check)",
        "6️⃣  Thống kê nhóm (Group By)",
        "7️⃣  Kiểm tra & báo cáo giá trị trống",
        "8️⃣  Gộp nhiều file thành 1",
        "9️⃣  Chuẩn hóa dữ liệu văn bản",
    ], key="ec_feature")

    st.divider()

    # ======== FEATURE 1: COMPARE TWO FILES ========
    if feature.startswith("1"):
        st.markdown("##### 🔍 So sánh 2 file — Tìm điểm khác biệt")
        if len(dfs) < 2:
            st.warning("⚠️ Cần tải lên ít nhất 2 file.")
            return

        file_keys = list(dfs.keys())
        c1, c2 = st.columns(2)
        with c1:
            src1 = st.selectbox("File gốc:", file_keys, key="cmp_src1")
        with c2:
            src2 = st.selectbox("File so sánh:", [k for k in file_keys if k != src1], key="cmp_src2")

        df1 = dfs[src1]
        df2 = dfs[src2]

        c3, c4 = st.columns(2)
        with c3:
            key_col1 = st.selectbox("Cột khóa (File 1):", df1.columns.tolist(), key="cmp_key1")
        with c4:
            key_col2 = st.selectbox("Cột khóa (File 2):", df2.columns.tolist(), key="cmp_key2")

        st.markdown("**Chọn cột cần so sánh nội dung (tùy chọn):**")
        c5, c6 = st.columns(2)
        with c5:
            cmp_cols1 = st.multiselect("Cột so sánh (File 1):", [c for c in df1.columns if c != key_col1], key="cmp_cols1")
        with c6:
            cmp_cols2 = st.multiselect("Cột so sánh (File 2):", [c for c in df2.columns if c != key_col2], key="cmp_cols2")

        if len(cmp_cols1) != len(cmp_cols2):
            st.warning("⚠️ Số lượng cột so sánh ở 2 file phải bằng nhau.")

        if st.button("🚀 BẮT ĐẦU SO SÁNH", type="primary", use_container_width=True, key="btn_cmp"):
            with st.spinner("Đang so sánh..."):
                res = feature_compare_two_files(df1, df2, key_col1, key_col2, cmp_cols1, cmp_cols2)

            st.markdown("##### 📊 Kết quả")
            stat = res["stats"]
            cols_s = st.columns(6)
            labels = list(stat.keys())
            vals = list(stat.values())
            colors = ["🔵", "🔵", "🟠", "🟢", "🟡", "🔴"]
            for i, col in enumerate(cols_s):
                col.metric(colors[i] + " " + labels[i], f"{vals[i]:,}")

            with st.expander(f"📋 Chỉ có trong {src1} ({len(res['only_in_file1'])} dòng)", expanded=len(res["only_in_file1"]) > 0):
                if not res["only_in_file1"].empty:
                    st.dataframe(res["only_in_file1"], use_container_width=True)
                    st.download_button("📥 Tải xuống", to_excel_bytes(res["only_in_file1"]),
                                       f"chi_co_trong_file1_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                       key="dl_only1")
                else:
                    st.success("Không có dòng nào chỉ xuất hiện trong file 1")

            with st.expander(f"📋 Chỉ có trong {src2} ({len(res['only_in_file2'])} dòng)", expanded=len(res["only_in_file2"]) > 0):
                if not res["only_in_file2"].empty:
                    st.dataframe(res["only_in_file2"], use_container_width=True)
                    st.download_button("📥 Tải xuống", to_excel_bytes(res["only_in_file2"]),
                                       f"chi_co_trong_file2_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                       key="dl_only2")
                else:
                    st.success("Không có dòng nào chỉ xuất hiện trong file 2")

            if not res["differences"].empty:
                with st.expander(f"⚠️ Dòng có sự khác biệt nội dung ({len(res['differences'])} dòng)", expanded=True):
                    st.dataframe(res["differences"], use_container_width=True)
                    st.download_button("📥 Tải xuống", to_excel_bytes(res["differences"]),
                                       f"khac_biet_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                       key="dl_diff")

            # Full report
            all_sheets = {
                f"Chi có trong {src1}": res["only_in_file1"] if not res["only_in_file1"].empty else pd.DataFrame({"Kết quả": ["Không có"]}),
                f"Chi có trong {src2}": res["only_in_file2"] if not res["only_in_file2"].empty else pd.DataFrame({"Kết quả": ["Không có"]}),
                "Khác biệt nội dung": res["differences"] if not res["differences"].empty else pd.DataFrame({"Kết quả": ["Không có khác biệt"]}),
                "Thống kê": pd.DataFrame([stat]).T.reset_index().rename(columns={"index": "Chỉ tiêu", 0: "Giá trị"}),
            }
            st.download_button("📥 Tải báo cáo đầy đủ (Excel)",
                               multi_sheet_excel(all_sheets),
                               f"BAO_CAO_SO_SANH_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                               key="dl_full_cmp")

    # ======== FEATURE 2: FIND DUPLICATES ========
    elif feature.startswith("2"):
        st.markdown("##### 🔁 Tìm & lọc dữ liệu TRÙNG")
        src = st.selectbox("Chọn file:", list(dfs.keys()), key="dup_src")
        df = dfs[src]

        dup_cols = st.multiselect("Cột xét trùng (chọn ≥1):", df.columns.tolist(), default=[df.columns[0]], key="dup_cols")
        keep_opt = st.radio("Giữ lại dòng nào khi xóa?", ["first — giữ dòng đầu tiên", "last — giữ dòng cuối", "none — xóa tất cả trùng"],
                            horizontal=True, key="dup_keep")
        keep_val = keep_opt.split(" ")[0] if keep_opt.startswith("first") or keep_opt.startswith("last") else False

        if st.button("🚀 TÌM DỮ LIỆU TRÙNG", type="primary", use_container_width=True, key="btn_dup"):
            if not dup_cols:
                st.warning("Chọn ít nhất 1 cột.")
                return
            with st.spinner("Đang xử lý..."):
                duplicates, cleaned, removed = feature_find_duplicates(df, dup_cols, keep=keep_val)

            m1, m2, m3 = st.columns(3)
            m1.metric("🔢 Tổng dòng gốc", f"{len(df):,}")
            m2.metric("🔴 Dòng trùng", f"{len(duplicates):,}")
            m3.metric("✅ Sau khi lọc", f"{len(cleaned):,}")

            with st.expander(f"🔴 Danh sách dữ liệu trùng ({len(duplicates)} dòng)", expanded=True):
                if not duplicates.empty:
                    st.dataframe(duplicates, use_container_width=True)
                    st.download_button("📥 Tải danh sách trùng",
                                       to_excel_bytes(duplicates, highlight_rows=list(range(len(duplicates)))),
                                       f"du_lieu_trung_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                       key="dl_dup")
                else:
                    st.success("✅ Không có dữ liệu trùng!")

            with st.expander(f"✅ File đã loại trùng ({len(cleaned)} dòng)"):
                st.dataframe(cleaned, use_container_width=True)
                st.download_button("📥 Tải file đã lọc trùng",
                                   to_excel_bytes(cleaned),
                                   f"da_loc_trung_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                   key="dl_cleaned")

    # ======== FEATURE 3: VLOOKUP ========
    elif feature.startswith("3"):
        st.markdown("##### 🔗 VLOOKUP — Ghép dữ liệu từ file tra cứu")
        if len(dfs) < 2:
            st.warning("⚠️ Cần tải lên ít nhất 2 file.")
            return

        file_keys = list(dfs.keys())
        c1, c2 = st.columns(2)
        with c1:
            main_src = st.selectbox("File chính (cần thêm dữ liệu):", file_keys, key="vl_main")
        with c2:
            lookup_src = st.selectbox("File tra cứu:", [k for k in file_keys if k != main_src], key="vl_lookup")

        df_main = dfs[main_src]
        df_lookup = dfs[lookup_src]

        c3, c4 = st.columns(2)
        with c3:
            main_key = st.selectbox("Cột khóa (File chính):", df_main.columns.tolist(), key="vl_mk")
        with c4:
            lookup_key = st.selectbox("Cột khóa (File tra cứu):", df_lookup.columns.tolist(), key="vl_lk")

        lookup_cols = st.multiselect("Chọn cột cần lấy từ file tra cứu:",
                                     [c for c in df_lookup.columns if c != lookup_key], key="vl_cols")

        if st.button("🚀 THỰC HIỆN VLOOKUP", type="primary", use_container_width=True, key="btn_vl"):
            if not lookup_cols:
                st.warning("Chọn ít nhất 1 cột để ghép.")
                return
            with st.spinner("Đang ghép dữ liệu..."):
                result, not_found = feature_vlookup(df_main, df_lookup, main_key, lookup_key, lookup_cols)

            st.metric("✅ Dòng ghép được", f"{len(df_main) - len(not_found):,} / {len(df_main):,}")
            if not_found:
                st.warning(f"⚠️ {len(not_found)} dòng không tìm thấy trong file tra cứu")
                with st.expander("Xem danh sách không khớp"):
                    st.write(not_found[:100])

            st.dataframe(result, use_container_width=True)
            st.download_button("📥 Tải kết quả VLOOKUP",
                               to_excel_bytes(result),
                               f"vlookup_result_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                               key="dl_vl")

    # ======== FEATURE 4: ADVANCED FILTER ========
    elif feature.startswith("4"):
        st.markdown("##### 🔍 Lọc dữ liệu nâng cao")
        src = st.selectbox("Chọn file:", list(dfs.keys()), key="flt_src")
        df = dfs[src]

        st.markdown("**Thiết lập điều kiện lọc:**")
        num_conds = st.number_input("Số điều kiện:", min_value=1, max_value=10, value=1, key="flt_num")

        operators = ["=", "≠", "chứa", "không chứa", "bắt đầu bằng", "trống", "không trống", ">", "<"]
        conditions = []
        for i in range(int(num_conds)):
            cc = st.columns([3, 2, 3])
            col_sel = cc[0].selectbox(f"Cột #{i+1}:", df.columns.tolist(), key=f"flt_col_{i}")
            op_sel = cc[1].selectbox("Toán tử:", operators, key=f"flt_op_{i}")
            val_sel = cc[2].text_input("Giá trị:", key=f"flt_val_{i}") if op_sel not in ["trống", "không trống"] else ""
            conditions.append({"col": col_sel, "operator": op_sel, "value": val_sel})

        if st.button("🚀 LỌC DỮ LIỆU", type="primary", use_container_width=True, key="btn_flt"):
            with st.spinner("Đang lọc..."):
                filtered = feature_filter_advanced(df, conditions)

            st.metric("📊 Kết quả", f"{len(filtered):,} / {len(df):,} dòng")
            st.dataframe(filtered, use_container_width=True)
            if not filtered.empty:
                st.download_button("📥 Tải kết quả lọc",
                                   to_excel_bytes(filtered),
                                   f"loc_du_lieu_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                   key="dl_flt")

    # ======== FEATURE 5: CROSS CHECK ========
    elif feature.startswith("5"):
        st.markdown("##### ✅ Kiểm tra chéo 2 cột (Cross-check)")
        if len(dfs) < 2:
            st.warning("⚠️ Cần tải lên ít nhất 2 file.")
            return

        file_keys = list(dfs.keys())
        c1, c2 = st.columns(2)
        with c1:
            src1 = st.selectbox("File 1:", file_keys, key="cc_src1")
        with c2:
            src2 = st.selectbox("File 2:", [k for k in file_keys if k != src1], key="cc_src2")

        df1 = dfs[src1]
        df2 = dfs[src2]

        cc1, cc2 = st.columns(2)
        with cc1:
            check_col1 = st.selectbox("Cột kiểm tra (File 1):", df1.columns.tolist(), key="cc_col1")
        with cc2:
            check_col2 = st.selectbox("Cột đối chiếu (File 2):", df2.columns.tolist(), key="cc_col2")

        if st.button("🚀 KIỂM TRA CHÉO", type="primary", use_container_width=True, key="btn_cc"):
            with st.spinner("Đang kiểm tra..."):
                missing_rows, miss2, miss1 = feature_cross_check(df1, df2, check_col1, check_col2)

            m1, m2 = st.columns(2)
            m1.metric(f"🔴 Trong {src1} nhưng KHÔNG có trong {src2}", f"{miss2:,}")
            m2.metric(f"🟠 Trong {src2} nhưng KHÔNG có trong {src1}", f"{miss1:,}")

            with st.expander(f"Dữ liệu trong {src1} không có trong {src2}:", expanded=True):
                if not missing_rows.empty:
                    st.dataframe(missing_rows, use_container_width=True)
                    st.download_button("📥 Tải xuống",
                                       to_excel_bytes(missing_rows),
                                       f"khong_co_trong_file2_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                       key="dl_cc")
                else:
                    st.success("✅ Tất cả dữ liệu đều khớp!")

    # ======== FEATURE 6: GROUP BY ========
    elif feature.startswith("6"):
        st.markdown("##### 📈 Thống kê nhóm (Group By)")
        src = st.selectbox("Chọn file:", list(dfs.keys()), key="grp_src")
        df = dfs[src]

        c1, c2 = st.columns(2)
        with c1:
            group_col = st.selectbox("Nhóm theo cột:", df.columns.tolist(), key="grp_col")
        with c2:
            agg_cols = st.multiselect("Cột thống kê (số liệu):", [c for c in df.columns if c != group_col], key="grp_agg")

        if st.button("🚀 THỐNG KÊ", type="primary", use_container_width=True, key="btn_grp"):
            if not agg_cols:
                st.warning("Chọn ít nhất 1 cột để thống kê.")
                return
            with st.spinner("Đang tính toán..."):
                result = feature_summary_stats(df, group_col, agg_cols)

            st.dataframe(result, use_container_width=True)
            st.download_button("📥 Tải bảng thống kê",
                               to_excel_bytes(result),
                               f"thong_ke_nhom_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                               key="dl_grp")

    # ======== FEATURE 7: MISSING VALUES ========
    elif feature.startswith("7"):
        st.markdown("##### 🔎 Kiểm tra & báo cáo giá trị trống")
        src = st.selectbox("Chọn file:", list(dfs.keys()), key="mv_src")
        df = dfs[src]

        if st.button("🚀 KIỂM TRA GIÁ TRỊ TRỐNG", type="primary", use_container_width=True, key="btn_mv"):
            with st.spinner("Đang kiểm tra..."):
                report = feature_check_missing_values(df)

            if report.empty:
                st.success("✅ Không có giá trị trống nào!")
            else:
                st.warning(f"⚠️ Tìm thấy {len(report)} cột có giá trị trống")
                st.dataframe(report, use_container_width=True)

                # Show rows with any empty
                problematic_cols = report["Cột"].tolist()
                mask = df[problematic_cols].apply(lambda s: (s.astype(str).str.strip() == "") | (s.astype(str).str.upper() == "NAN")).any(axis=1)
                df_with_empty = df[mask].copy()
                st.markdown(f"**Các dòng có giá trị trống ({len(df_with_empty)} dòng):**")
                st.dataframe(df_with_empty, use_container_width=True)

                sheets = {
                    "Báo cáo cột trống": report,
                    "Dòng có giá trị trống": df_with_empty
                }
                st.download_button("📥 Tải báo cáo giá trị trống",
                                   multi_sheet_excel(sheets),
                                   f"bao_cao_trong_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                   key="dl_mv")

    # ======== FEATURE 8: CONCAT FILES ========
    elif feature.startswith("8"):
        st.markdown("##### 📎 Gộp nhiều file thành 1")
        selected_files = st.multiselect("Chọn file cần gộp:", list(dfs.keys()), default=list(dfs.keys()), key="cat_files")
        add_source = st.checkbox("Thêm cột '__Nguồn__' để biết dòng đến từ file nào", value=True, key="cat_src")

        if st.button("🚀 GỘP FILE", type="primary", use_container_width=True, key="btn_cat"):
            if len(selected_files) < 2:
                st.warning("Chọn ít nhất 2 file để gộp.")
                return
            with st.spinner("Đang gộp..."):
                selected_dfs = [dfs[k] for k in selected_files]
                result = feature_concat_files(selected_dfs, add_source=add_source, source_names=selected_files)

            m1, m2 = st.columns(2)
            m1.metric("📄 Tổng dòng sau gộp", f"{len(result):,}")
            m2.metric("📋 Số cột", f"{len(result.columns)}")

            st.dataframe(result, use_container_width=True)
            st.download_button("📥 Tải file đã gộp",
                               to_excel_bytes(result),
                               f"gop_file_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                               key="dl_cat")

    # ======== FEATURE 9: NORMALIZE ========
    elif feature.startswith("9"):
        st.markdown("##### 🧹 Chuẩn hóa dữ liệu văn bản")
        src = st.selectbox("Chọn file:", list(dfs.keys()), key="norm_src")
        df = dfs[src]

        norm_cols = st.multiselect("Cột cần chuẩn hóa (trim khoảng trắng):", df.columns.tolist(), key="norm_cols")
        st.caption("Thao tác: Xóa khoảng trắng đầu/cuối (trim)")

        if st.button("🚀 CHUẨN HÓA DỮ LIỆU", type="primary", use_container_width=True, key="btn_norm"):
            if not norm_cols:
                st.warning("Chọn ít nhất 1 cột.")
                return
            with st.spinner("Đang chuẩn hóa..."):
                result = feature_normalize_data(df, norm_cols)

            st.success(f"✅ Đã chuẩn hóa {len(norm_cols)} cột với {len(result):,} dòng")
            st.dataframe(result, use_container_width=True)
            st.download_button("📥 Tải file đã chuẩn hóa",
                               to_excel_bytes(result),
                               f"chuan_hoa_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                               key="dl_norm")
