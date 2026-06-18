# -*- coding: utf-8 -*-
"""
bhyt_rules_bo_sung.py
File chứa các ĐIỀU KIỆN ĐỐI CHIẾU BỔ SUNG, tách riêng khỏi bhyt_rules.py để
khi cần thêm/sửa điều kiện mới trong tương lai sẽ không ảnh hưởng tới cấu
trúc 4 nhóm quy tắc gốc (Hành chính, Danh mục-Giá, Hợp lý chỉ định, Thời gian).

Mỗi hàm check_* nhận vào dữ liệu cần thiết và trả về list các dict lỗi có
cùng cấu trúc với bhyt_rules._err (ma_lk, nhom, muc_do, truong_loi, gia_tri,
mo_ta_loi, huong_xu_ly, nguon_file) để tương thích với cơ chế gộp báo cáo
hiện tại trong bhyt_app.py.

CÁC ĐIỀU KIỆN ĐÃ BỔ SUNG (theo yêu cầu cập nhật):
  1. Người thực hiện DVKT (XML3, NGUOI_THUC_HIEN) được mở rộng chấp nhận
     thêm CHUCDANH_NN = 2 (Y sĩ), ngoài 6 và 10 (kỹ thuật viên) đã có.
  2. DVKT có TEN_DICH_VU bắt đầu bằng "Khám" (không phân biệt hoa/thường)
     -> NGUOI_THUC_HIEN bắt buộc phải là Bác sĩ (CHUCDANH_NN = 1), không
     chấp nhận các chức danh khác trong trường hợp này.
  3. Trường hợp đặc biệt: mã chứng chỉ hành nghề '0870/CCHN-D-SYT-TNI' luôn
     được phép thực hiện DVKT thuộc nhóm X-quang, bỏ qua mọi kiểm tra
     CHUCDANH_NN khác (ngoại lệ riêng cho người này).

Ghi chú thiết kế: 2 điều kiện (2) và (3) được ưu tiên kiểm tra trước điều
kiện (1) - tức nếu DVKT là "Khám" thì áp dụng luôn yêu cầu Bác sĩ, không
còn xét tới {2,6,10} hay phân công nữa; nếu là X-quang và đúng mã CCHN đặc
biệt thì luôn đạt, không cần xét gì thêm.
"""

import re

from bhyt_rules import _text, _err, _tra_nhan_vien, CHUCDANH_BAC_SI

# Chức danh nghề nghiệp được phép thực hiện DVKT (mở rộng so với bản gốc):
# 2 = Y sĩ, 6 và 10 = kỹ thuật viên (giữ như cũ)
CHUCDANH_THUC_HIEN_DVKT_BO_SUNG = {"2", "6", "10"}

# Mã chứng chỉ hành nghề được phép thực hiện MỌI DVKT X-quang, không cần
# xét CHUCDANH_NN hay phân công (trường hợp đặc biệt do đơn vị xác nhận)
MACCHN_DAC_BIET_XQUANG = "0870/CCHN-D-SYT-TNI"

# Từ khóa nhận diện DVKT là X-quang (dùng để áp dụng ngoại lệ MACCHN_DAC_BIET_XQUANG)
TU_KHOA_XQUANG = ["X-QUANG", "XQUANG", "X QUANG"]


def _la_dvkt_kham(ten_dv):
    """True nếu tên dịch vụ bắt đầu bằng từ 'Khám' (không phân biệt hoa/thường,
    cho phép có khoảng trắng ở đầu)."""
    if not ten_dv:
        return False
    return bool(re.match(r"^\s*khám\b", ten_dv.strip(), flags=re.IGNORECASE))


def _la_dvkt_xquang(ten_dv):
    """True nếu tên dịch vụ thuộc nhóm X-quang."""
    if not ten_dv:
        return False
    ten_upper = ten_dv.upper()
    return any(k in ten_upper for k in TU_KHOA_XQUANG)


def check_nguoi_thuc_hien_dvkt_bo_sung(ho_so, ma_lk, ma_hsba, ho_ten, nhan_vien_yte):
    """
    Kiểm tra người thực hiện DVKT (XML3, trường NGUOI_THUC_HIEN) theo các
    điều kiện bổ sung mới nhất:

      a) Nếu DVKT là "Khám..." -> NGUOI_THUC_HIEN PHẢI là Bác sĩ (CDNN=1).
         Không xét các điều kiện khác trong trường hợp này.
      b) Nếu DVKT là X-quang VÀ NGUOI_THUC_HIEN = mã CCHN đặc biệt
         '0870/CCHN-D-SYT-TNI' -> LUÔN ĐẠT, không cần kiểm tra gì thêm.
      c) Các trường hợp còn lại: đạt khi CHUCDANH_NN thuộc {2,6,10}
         HOẶC có DVKT_KHAC + VB_PHANCONG phù hợp (hai điều kiện độc lập - OR).

    ho_so: phần tử XML (ET.Element) của 1 dòng HO_SO_KCB trong XML3.
    ma_lk: mã liên kết hồ sơ (dùng để gắn lỗi).
    ma_hsba, ho_ten: thông tin tra cứu sẵn từ XML1 (Mã hồ sơ bệnh án, Họ tên
        người bệnh) - LUÔN đính kèm vào mọi lỗi xuất ra theo yêu cầu hiển thị
        đầy đủ Tên + Mã hồ sơ bệnh án.
    nhan_vien_yte: dict danh mục nhân viên y tế {MACCHN: {...}} (từ
        bhyt_rules.load_default_danh_muc()).

    Trả về list dict lỗi.
    """
    errors = []
    nguon = "XML3 - DVKT (điều kiện bổ sung)"

    ma_dv = _text(ho_so, "MA_DICH_VU")
    ten_dv = _text(ho_so, "TEN_DICH_VU")
    ma_vat_tu = _text(ho_so, "MA_VAT_TU")
    nguoi_th = _text(ho_so, "NGUOI_THUC_HIEN")

    la_giuong = "giường" in (ten_dv or "").lower() or "giuong" in (ten_dv or "").lower()
    # VTYT và giường bệnh không có người thực hiện chuyên môn -> không áp dụng quy tắc này
    if ma_vat_tu or la_giuong or not nguoi_th:
        return errors

    macchn = nguoi_th.strip()

    def _gan_thong_tin_hsba(d):
        """Đính kèm Mã HSBA + Họ tên vào dict lỗi (yêu cầu hiển thị đầy đủ)."""
        d["ma_hsba"] = ma_hsba or ""
        d["ho_ten"] = ho_ten or ""
        return d

    # ---- b) Ngoại lệ: mã CCHN đặc biệt luôn được phép thực hiện X-quang ----
    if _la_dvkt_xquang(ten_dv) and macchn == MACCHN_DAC_BIET_XQUANG:
        return errors  # luôn đạt, không kiểm tra gì thêm

    nv = _tra_nhan_vien(macchn, nhan_vien_yte)

    # ---- a) DVKT "Khám..." -> bắt buộc Bác sĩ (CDNN=1) ----
    if _la_dvkt_kham(ten_dv):
        if nv is None:
            if nhan_vien_yte:
                errors.append(_gan_thong_tin_hsba(_err(
                    ma_lk, "Điều kiện bổ sung - Nhân sự", "Cảnh báo",
                    "NGUOI_THUC_HIEN", macchn,
                    f"Không tìm thấy mã chứng chỉ hành nghề '{macchn}' (người thực hiện công khám "
                    f"'{ma_dv}' - {ten_dv}) trong danh mục nhân viên y tế.",
                    "Kiểm tra lại NGUOI_THUC_HIEN hoặc bổ sung nhân viên vào danh mục nhân viên y tế.",
                    nguon
                )))
            return errors
        cd = nv.get("chucdanh", "")
        if cd != CHUCDANH_BAC_SI:
            ten_cd = {"2": "Y sĩ", "3": "Điều dưỡng", "4": "Dược sĩ"}.get(cd, f"Chức danh {cd}")
            errors.append(_gan_thong_tin_hsba(_err(
                ma_lk, "Điều kiện bổ sung - Nhân sự", "Lỗi",
                "NGUOI_THUC_HIEN", macchn,
                f"Dịch vụ '{ma_dv}' ({ten_dv}) là công khám, người thực hiện là "
                f"{nv.get('ho_ten', '')} ({ten_cd}) - không phải Bác sĩ (CHUCDANH_NN=1). "
                "Công khám bắt buộc phải do Bác sĩ thực hiện.",
                "Kiểm tra lại người thực hiện (NGUOI_THUC_HIEN); công khám phải do Bác sĩ thực hiện, "
                "không áp dụng ngoại lệ phân công cho trường hợp này.",
                nguon
            )))
        return errors

    # ---- c) Trường hợp chung: {2,6,10} HOẶC phân công riêng (OR độc lập) ----
    if nv is None:
        if nhan_vien_yte:
            errors.append(_gan_thong_tin_hsba(_err(
                ma_lk, "Điều kiện bổ sung - Nhân sự", "Cảnh báo",
                "NGUOI_THUC_HIEN", macchn,
                f"Không tìm thấy mã chứng chỉ hành nghề '{macchn}' (người thực hiện DVKT "
                f"'{ma_dv}' - {ten_dv}) trong danh mục nhân viên y tế.",
                "Kiểm tra lại NGUOI_THUC_HIEN hoặc bổ sung nhân viên vào danh mục nhân viên y tế.",
                nguon
            )))
        return errors

    cd = nv.get("chucdanh", "")
    dat_theo_chucdanh = cd in CHUCDANH_THUC_HIEN_DVKT_BO_SUNG
    dat_theo_phancong = bool(nv.get("dvkt_khac")) and bool(nv.get("vb_phancong")) and \
        (not ma_dv or ma_dv.strip() in (nv.get("dvkt_khac") or ""))

    if dat_theo_chucdanh or dat_theo_phancong:
        return errors  # đạt yêu cầu - OR độc lập, không xuất lỗi

    ten_cd = {"3": "Điều dưỡng", "4": "Dược sĩ", "1": "Bác sĩ"}.get(cd, f"Chức danh {cd}")
    errors.append(_gan_thong_tin_hsba(_err(
        ma_lk, "Điều kiện bổ sung - Nhân sự", "Cảnh báo",
        "NGUOI_THUC_HIEN", macchn,
        f"Người thực hiện DVKT '{ma_dv}' ({ten_dv}) - {nv.get('ho_ten', '')} ({ten_cd}) - "
        "không thuộc chức danh được phép (CHUCDANH_NN=2,6,10) và không có văn bản phân công "
        "thực hiện kỹ thuật khác phù hợp.",
        "Kiểm tra lại người thực hiện hoặc bổ sung văn bản phân công (DVKT_KHAC/VB_PHANCONG) "
        "trong danh mục nhân viên y tế nếu việc phân công là có thật.",
        nguon
    )))
    return errors


def check_dvkt_bo_sung_full(ho_so, ma_lk, ma_hsba, ho_ten, danh_muc):
    """
    Hàm tổng hợp - gọi tất cả các quy tắc bổ sung áp dụng cho 1 dòng XML3.
    Dùng hàm này trong bhyt_app.py để dễ mở rộng thêm quy tắc khác trong
    tương lai mà không cần sửa lời gọi ở bhyt_app.py.
    """
    nhan_vien_yte = (danh_muc or {}).get("nhan_vien_yte", {})
    errors = []
    errors.extend(check_nguoi_thuc_hien_dvkt_bo_sung(ho_so, ma_lk, ma_hsba, ho_ten, nhan_vien_yte))
    return errors
