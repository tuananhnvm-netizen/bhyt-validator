# -*- coding: utf-8 -*-
"""
bhyt_rules.py
Bộ quy tắc kiểm tra dữ liệu XML1-XML5 (Quyết định 4210/QĐ-BYT) trước khi đẩy
lên Cổng tiếp nhận / Giám định BHYT.

4 nhóm quy tắc:
  1. Hành chính & thẻ BHYT
  2. Danh mục thuốc / DVKT / VTYT, giá, định mức kinh tế kỹ thuật
  3. Hợp lý chỉ định y khoa (bắt chéo ICD-10, giới tính/tuổi, trùng lặp, liều dùng)
  4. Thời gian KCB & trùng lặp đợt điều trị

Mỗi lỗi trả về là 1 dict với khóa thống nhất:
    ma_lk        : Mã liên kết hồ sơ
    nhom         : Tên nhóm quy tắc (1..4)
    muc_do       : "Lỗi" (chặn, chắc chắn bị xuất toán/từ chối) hoặc
                   "Cảnh báo" (cần bác sĩ/điều dưỡng xem lại)
    truong_loi   : Tên trường dữ liệu liên quan
    gia_tri      : Giá trị hiện tại trong hồ sơ (nếu có)
    mo_ta_loi    : Diễn giải lỗi
    huong_xu_ly  : Gợi ý cách sửa
    nguon_file   : Loại file (XML1..XML5)
"""

from datetime import datetime
import re

# =====================================================================
#  HẰNG SỐ DÙNG CHUNG
# =====================================================================

# Mã quyền lợi hưởng -> tỷ lệ hưởng BHYT (%) - theo Luật BHYT hiện hành
# (mã 5 - một số đối tượng đặc thù 100% có điều kiện - cần đối chiếu thêm
#  đối tượng tham gia, ở đây chỉ dùng cảnh báo nếu mã ngoài danh sách)
MA_MUC_HUONG_TY_LE = {
    "1": 100,
    "2": 95,
    "3": 80,
    "4": 100,   # nhóm ưu tiên đặc biệt (theo đối tượng)
    "5": 100,   # nhóm ưu tiên đặc biệt (theo đối tượng)
}

# Định dạng ngày giờ chuẩn 4210: YYYYMMDDHHMM (12 ký tự)
DATE_FMT = "%Y%m%d%H%M"
DATE_RE = re.compile(r"^\d{12}$")

# Các DVKT/CĐHA "nặng" thường bị soi khi không có ICD phù hợp
# -> dùng để cảnh báo khi KHÔNG có bảng đối chiếu ICD-DVKT do người dùng tải lên
HIGH_COST_DVKT_KEYWORDS = [
    "CT", "MRI", "CỘNG HƯỞNG TỪ", "CẮT LỚP VI TÍNH", "NỘI SOI",
    "SIÊU ÂM", "ĐIỆN TÂM ĐỒ", "XQUANG", "X-QUANG", "X QUANG",
]

# Từ khóa thuốc/DVKT chỉ dành riêng cho một giới tính (mẫu cơ bản, mở rộng
# thêm qua danh mục do người dùng tải lên nếu cần)
FEMALE_ONLY_KEYWORDS = [
    "THAI", "SẢN", "ÂM ĐẠO", "TỬ CUNG", "BUỒNG TRỨNG", "VÒNG TRÁNH THAI",
    "PHỤ KHOA", "CỔ TỬ CUNG", "VÚ", "TUYẾN VÚ", "MANG THAI",
]
MALE_ONLY_KEYWORDS = [
    "TUYẾN TIỀN LIỆT", "TIỀN LIỆT TUYẾN", "DƯƠNG VẬT", "TINH HOÀN",
    "BAO QUY ĐẦU", "NAM KHOA",
]

# Từ khóa DVKT/thuốc chỉ dành cho trẻ em (theo tuổi < 16, mở rộng tuỳ danh mục)
PEDIATRIC_KEYWORDS = ["NHI", "SƠ SINH", "TRẺ EM"]


# =====================================================================
#  HÀM TIỆN ÍCH
# =====================================================================

def _text(ho_so, tag):
    """Lấy text đã strip của 1 tag con, trả '' nếu không có."""
    el = ho_so.find(tag)
    return (el.text or "").strip() if el is not None and el.text else ""


def _to_float(value):
    try:
        return float(str(value).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def _to_datetime(value):
    value = (value or "").strip()
    if not DATE_RE.match(value):
        return None
    try:
        return datetime.strptime(value, DATE_FMT)
    except ValueError:
        return None


def _err(ma_lk, nhom, muc_do, truong_loi, gia_tri, mo_ta_loi, huong_xu_ly, nguon_file):
    return {
        "ma_lk": ma_lk or "Không xác định",
        "nhom": nhom,
        "muc_do": muc_do,
        "truong_loi": truong_loi,
        "gia_tri": gia_tri,
        "mo_ta_loi": mo_ta_loi,
        "huong_xu_ly": huong_xu_ly,
        "nguon_file": nguon_file,
    }


# =====================================================================
#  NHÓM 1: HÀNH CHÍNH & THẺ BHYT  (chạy trên dữ liệu XML1 - Hành chính)
# =====================================================================

def check_nhom1_hanh_chinh(ho_so, ma_lk, danh_muc=None):
    """
    Kiểm tra:
      1.1 Hiệu lực thẻ BHYT (GT_THE_TU / GT_THE_DEN so với NGAY_VAO / NGAY_RA)
      1.2 Khớp định danh Họ tên / Ngày sinh / Giới tính (đã chuẩn hoá cơ bản)
      1.3 Mã quyền lợi (MA_LOAI_KCB / MA_DOITUONG_KCB -> MUC_HUONG) hợp lệ và
          khớp với tỷ lệ hưởng đã khai (TYLE_BHTT / MUC_HUONG)
      1.4 Thông tin chuyển tuyến (MA_LOAI_RV / NOICHUYEN / SO_GIAY_CHUYEN_TUYEN)
    """
    errors = []
    nguon = "XML1 - Hành chính"

    # ---- 1.1 Hiệu lực thẻ BHYT ----
    gt_the_tu = _text(ho_so, "GT_THE_TU")
    gt_the_den = _text(ho_so, "GT_THE_DEN")
    ngay_vao = _text(ho_so, "NGAY_VAO")
    ngay_ra = _text(ho_so, "NGAY_RA")

    dt_tu = _to_datetime(gt_the_tu) if gt_the_tu else None
    dt_den = _to_datetime(gt_the_den) if gt_the_den else None
    dt_vao = _to_datetime(ngay_vao)
    dt_ra = _to_datetime(ngay_ra)

    if gt_the_tu and gt_the_den:
        if dt_tu is None or dt_den is None:
            errors.append(_err(
                ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                "GT_THE_TU/GT_THE_DEN", f"{gt_the_tu} - {gt_the_den}",
                "Định dạng giá trị thẻ (GT_THE_TU/GT_THE_DEN) không đúng YYYYMMDDHHMM.",
                "Kiểm tra lại định dạng ngày trên thẻ BHYT (12 số: năm/tháng/ngày/giờ/phút).", nguon
            ))
        else:
            if dt_vao and dt_vao < dt_tu:
                errors.append(_err(
                    ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                    "NGAY_VAO/GT_THE_TU", f"Vào viện: {ngay_vao} | Thẻ có giá trị từ: {gt_the_tu}",
                    "Ngày vào viện trước thời điểm thẻ BHYT có giá trị sử dụng -> thẻ chưa có hiệu lực.",
                    "Kiểm tra lại ngày vào viện hoặc kiểm tra thông tin gia hạn thẻ BHYT của người bệnh.", nguon
                ))
            if dt_ra and dt_den and dt_ra > dt_den:
                errors.append(_err(
                    ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                    "NGAY_RA/GT_THE_DEN", f"Ra viện: {ngay_ra} | Thẻ có giá trị đến: {gt_the_den}",
                    "Ngày ra viện sau thời điểm thẻ BHYT hết hạn -> phần chi phí phát sinh sau "
                    "ngày hết hạn thẻ có thể không được thanh toán.",
                    "Đề nghị người bệnh gia hạn/đóng tiếp BHYT hoặc kiểm tra lại GT_THE_DEN; "
                    "nếu thẻ đã được gia hạn cần cập nhật lại GT_THE_DEN cho đúng.", nguon
                ))
    else:
        errors.append(_err(
            ma_lk, "1. Hành chính & thẻ BHYT", "Cảnh báo",
            "GT_THE_TU/GT_THE_DEN", "(thiếu)",
            "Thiếu thông tin giá trị sử dụng thẻ BHYT (GT_THE_TU/GT_THE_DEN) để kiểm tra hiệu lực thẻ.",
            "Bổ sung GT_THE_TU, GT_THE_DEN lấy từ kết quả tra cứu thẻ BHYT (Cổng tiếp nhận).", nguon
        ))

    # ---- 1.2 Khớp định danh: Họ tên / Ngày sinh / Giới tính ----
    ho_ten = _text(ho_so, "HO_TEN")
    ngay_sinh = _text(ho_so, "NGAY_SINH")
    gioi_tinh = _text(ho_so, "GIOI_TINH")

    if ho_ten:
        # Cảnh báo các ký tự bất thường (số, ký tự đặc biệt) trong họ tên
        if re.search(r"[0-9@#$%^&*()_+=\[\]{};:\"\\|<>/~`]", ho_ten):
            errors.append(_err(
                ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                "HO_TEN", ho_ten,
                "Họ tên chứa ký tự số hoặc ký tự đặc biệt, không hợp lệ.",
                "Sửa lại Họ tên đúng theo CCCD/Thẻ BHYT, chỉ gồm chữ cái và dấu cách.", nguon
            ))
        if ho_ten != ho_ten.strip() or "  " in ho_ten:
            errors.append(_err(
                ma_lk, "1. Hành chính & thẻ BHYT", "Cảnh báo",
                "HO_TEN", ho_ten,
                "Họ tên có khoảng trắng thừa ở đầu/cuối hoặc giữa các từ.",
                "Chuẩn hoá lại Họ tên: xoá khoảng trắng thừa, viết hoa chữ đầu mỗi từ.", nguon
            ))

    if ngay_sinh:
        # Cho phép YYYYMMDDHHMM hoặc YYYY (chỉ năm sinh)
        if not (re.match(r"^\d{12}$", ngay_sinh) or re.match(r"^\d{4}$", ngay_sinh)):
            errors.append(_err(
                ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                "NGAY_SINH", ngay_sinh,
                "Ngày sinh không đúng định dạng (YYYYMMDDHHMM hoặc YYYY nếu chỉ có năm sinh).",
                "Sửa lại Ngày sinh đúng định dạng theo CSDL quốc gia về bảo hiểm.", nguon
            ))
        else:
            if re.match(r"^\d{12}$", ngay_sinh):
                dt_sinh = _to_datetime(ngay_sinh)
                if dt_sinh and dt_vao and dt_sinh > dt_vao:
                    errors.append(_err(
                        ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                        "NGAY_SINH", ngay_sinh,
                        "Ngày sinh lớn hơn (sau) ngày vào viện - không hợp lý.",
                        "Kiểm tra lại Ngày sinh của người bệnh theo CCCD/giấy khai sinh.", nguon
                    ))

    if gioi_tinh:
        if gioi_tinh not in ("1", "2", "3"):
            errors.append(_err(
                ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                "GIOI_TINH", gioi_tinh,
                "Mã giới tính không hợp lệ (chỉ nhận 1=Nam, 2=Nữ, 3=Khác/không xác định).",
                "Sửa lại GIOI_TINH theo đúng mã quy định (1: Nam, 2: Nữ).", nguon
            ))

    # ---- 1.3 Mã đối tượng và mức hưởng ----
    ma_doi_tuong = _text(ho_so, "MA_DOITUONG_KCB")
    muc_huong = _text(ho_so, "MUC_HUONG")
    ty_le_bhtt = _text(ho_so, "TYLE_BHTT")

    if muc_huong:
        if muc_huong not in MA_MUC_HUONG_TY_LE:
            errors.append(_err(
                ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                "MUC_HUONG", muc_huong,
                "Mã mức hưởng (MUC_HUONG) không thuộc danh mục hợp lệ (1-5).",
                "Tra lại mã mức hưởng trên thẻ BHYT/kết quả kiểm tra thẻ, sửa MUC_HUONG cho đúng.", nguon
            ))
        else:
            ty_le_so = _to_float(ty_le_bhtt)
            ty_le_chuan = MA_MUC_HUONG_TY_LE[muc_huong]
            if ty_le_so is not None and abs(ty_le_so - ty_le_chuan) > 0.01:
                if muc_huong in ("1", "2", "3"):
                    errors.append(_err(
                        ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                        "TYLE_BHTT", ty_le_bhtt,
                        f"Tỷ lệ BHTT khai ({ty_le_bhtt}%) không khớp với tỷ lệ chuẩn của mức hưởng "
                        f"MUC_HUONG={muc_huong} (chuẩn = {ty_le_chuan}%).",
                        "Kiểm tra lại MUC_HUONG và TYLE_BHTT, đảm bảo khớp đúng theo mức hưởng trên thẻ.", nguon
                    ))
                else:
                    errors.append(_err(
                        ma_lk, "1. Hành chính & thẻ BHYT", "Cảnh báo",
                        "TYLE_BHTT", ty_le_bhtt,
                        f"Tỷ lệ BHTT khai ({ty_le_bhtt}%) khác mặc định ({ty_le_chuan}%) đối với "
                        f"MUC_HUONG={muc_huong}.",
                        "Mã mức hưởng 4/5 có thể có tỷ lệ đặc biệt theo từng đối tượng ưu tiên "
                        "(người có công, trẻ em dưới 6 tuổi, hộ nghèo...). Đối chiếu với quyết định/"
                        "giấy chứng nhận đối tượng để xác nhận TYLE_BHTT đã đúng trước khi gửi.", nguon
                    ))
    else:
        errors.append(_err(
            ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
            "MUC_HUONG", "(thiếu)",
            "Thiếu mã mức hưởng (MUC_HUONG) - không thể tính tỷ lệ thanh toán BHYT.",
            "Bổ sung MUC_HUONG (1-5) lấy từ kết quả tra cứu thẻ BHYT.", nguon
        ))

    # ---- 1.4 Thông tin chuyển tuyến ----
    # Dấu hiệu chuyển tuyến: chỉ dựa vào MA_NOI_DI / SO_GIAYCHUYEN đã có giá trị
    # (không dựa vào MA_LOAI_KCB vì cách phân loại mã này có thể khác nhau
    # giữa các phiên bản danh mục dùng chung).
    ma_noi_di = _text(ho_so, "MA_NOI_DI")
    so_giay_chuyen_tuyen = _text(ho_so, "SO_GIAYCHUYEN")

    if ma_noi_di or so_giay_chuyen_tuyen:
        # Có dấu hiệu chuyển tuyến -> bắt buộc đủ cả mã nơi chuyển đi và số giấy chuyển tuyến
        if not ma_noi_di:
            errors.append(_err(
                ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                "MA_NOI_DI", "(thiếu)",
                "Hồ sơ có Số giấy chuyển tuyến (SO_GIAYCHUYEN) nhưng thiếu mã cơ sở chuyển đi (MA_NOI_DI).",
                "Bổ sung mã cơ sở y tế chuyển đi (MA_NOI_DI) theo Giấy chuyển tuyến.", nguon
            ))
        if not so_giay_chuyen_tuyen:
            errors.append(_err(
                ma_lk, "1. Hành chính & thẻ BHYT", "Lỗi",
                "SO_GIAYCHUYEN", "(thiếu)",
                "Hồ sơ có mã cơ sở chuyển đi (MA_NOI_DI) nhưng thiếu Số giấy chuyển tuyến (SO_GIAYCHUYEN).",
                "Bổ sung số Giấy chuyển tuyến/Phiếu chuyển tuyến đầy đủ, đúng số đã cấp tại nơi chuyển đi.", nguon
            ))

    return errors


# =====================================================================
#  NHÓM 2: DANH MỤC THUỐC / DVKT / VTYT, GIÁ, ĐỊNH MỨC KINH TẾ KỸ THUẬT
# =====================================================================

def check_nhom2_danh_muc_gia(ho_so, ma_lk, file_type, danh_muc=None):
    """
    danh_muc (optional, do người dùng tải lên) có thể chứa:
      danh_muc['gia_thuoc']   : dict {MA_THUOC: gia_duyet}
      danh_muc['gia_dvkt']    : dict {MA_DICH_VU: gia_duyet}
      danh_muc['gia_vtyt']    : dict {MA_VTYT: gia_duyet}
      danh_muc['dinh_muc']    : dict {MA_DICH_VU: {MA_VTYT: so_luong_toi_da}}
    Nếu không có danh mục, chỉ kiểm tra các quy tắc logic nội bộ
    (đơn giá x số lượng = thành tiền, thành tiền BH <= thành tiền, ...).
    """
    errors = []
    danh_muc = danh_muc or {}

    if file_type == "Thuốc":
        nguon = "XML2 - Thuốc"
        ma_thuoc = _text(ho_so, "MA_THUOC")
        ten_thuoc = _text(ho_so, "TEN_THUOC")
        so_luong = _to_float(_text(ho_so, "SO_LUONG"))
        don_gia = _to_float(_text(ho_so, "DON_GIA"))
        thanh_tien_bh = _to_float(_text(ho_so, "THANH_TIEN_BH"))
        thanh_tien_bn_tt = _to_float(_text(ho_so, "T_TONGCHI"))

        # 2.1 Đối chiếu giá trúng thầu/giá duyệt
        gia_thuoc_dm = danh_muc.get("gia_thuoc", {})
        if gia_thuoc_dm:
            if ma_thuoc and ma_thuoc not in gia_thuoc_dm:
                errors.append(_err(
                    ma_lk, "2. Danh mục - Giá - Định mức", "Lỗi",
                    "MA_THUOC", ma_thuoc,
                    f"Mã thuốc '{ma_thuoc}' ({ten_thuoc}) không có trong danh mục thuốc đã trúng thầu/"
                    "được phê duyệt sử dụng tại cơ sở.",
                    "Kiểm tra lại mã thuốc theo danh mục trúng thầu hiện hành tại cơ sở, "
                    "hoặc bổ sung thuốc vào danh mục nếu đã có quyết định phê duyệt bổ sung.", nguon
                ))
            elif ma_thuoc and don_gia is not None:
                gia_duyet = gia_thuoc_dm.get(ma_thuoc)
                if gia_duyet is not None and abs(don_gia - gia_duyet) > 0.5:
                    errors.append(_err(
                        ma_lk, "2. Danh mục - Giá - Định mức", "Lỗi",
                        "DON_GIA", don_gia,
                        f"Đơn giá thuốc '{ma_thuoc}' ({ten_thuoc}) = {don_gia:,.0f} không khớp giá "
                        f"trúng thầu được duyệt = {gia_duyet:,.0f}.",
                        "Sửa lại DON_GIA theo đúng giá trúng thầu/giá duyệt hiện hành tại cơ sở.", nguon
                    ))

        # 2.2 Kiểm tra logic số học: đơn giá x số lượng (cho phép sai số làm tròn)
        if so_luong is not None and don_gia is not None and thanh_tien_bh is not None:
            tinh_lai = round(so_luong * don_gia, 0)
            if abs(tinh_lai - round(thanh_tien_bh + (thanh_tien_bn_tt or 0) -
                                     (thanh_tien_bn_tt or 0), 0)) > max(1, 0.01 * tinh_lai) \
               and thanh_tien_bn_tt is None:
                # Khi không có cột tổng chi để đối chiếu phần BN tự trả,
                # chỉ kiểm tra thành tiền BH không vượt SL*đơn giá
                pass
            if thanh_tien_bh is not None and tinh_lai is not None and thanh_tien_bh - tinh_lai > max(1, 0.01 * tinh_lai):
                errors.append(_err(
                    ma_lk, "2. Danh mục - Giá - Định mức", "Lỗi",
                    "THANH_TIEN_BH", thanh_tien_bh,
                    f"Thành tiền BHYT ({thanh_tien_bh:,.0f}) lớn hơn Số lượng x Đơn giá "
                    f"({so_luong:g} x {don_gia:,.0f} = {tinh_lai:,.0f}).",
                    "Kiểm tra lại công thức tính thành tiền BHYT, đảm bảo "
                    "Thành tiền BH <= Số lượng x Đơn giá.", nguon
                ))

        if so_luong is not None and so_luong <= 0:
            errors.append(_err(
                ma_lk, "2. Danh mục - Giá - Định mức", "Lỗi",
                "SO_LUONG", so_luong,
                "Số lượng thuốc phải lớn hơn 0.",
                "Kiểm tra lại số lượng thuốc đã cấp/sử dụng cho người bệnh.", nguon
            ))

    elif file_type == "Dịch vụ kỹ thuật":
        nguon = "XML3 - DVKT"
        ma_dv = _text(ho_so, "MA_DICH_VU")
        ten_dv = _text(ho_so, "TEN_DICH_VU")
        so_luong = _to_float(_text(ho_so, "SO_LUONG"))
        don_gia_bh = _to_float(_text(ho_so, "DON_GIA_BH"))
        thanh_tien_bh = _to_float(_text(ho_so, "THANH_TIEN_BH"))

        gia_dvkt_dm = danh_muc.get("gia_dvkt", {})
        if gia_dvkt_dm:
            if ma_dv and ma_dv not in gia_dvkt_dm:
                errors.append(_err(
                    ma_lk, "2. Danh mục - Giá - Định mức", "Lỗi",
                    "MA_DICH_VU", ma_dv,
                    f"Mã DVKT '{ma_dv}' ({ten_dv}) không có trong danh mục kỹ thuật được phê duyệt "
                    "thực hiện tại cơ sở.",
                    "Kiểm tra lại mã DVKT theo danh mục kỹ thuật đã được phê duyệt phân loại phẫu thuật, "
                    "thủ thuật/Danh mục dùng chung.", nguon
                ))
            elif ma_dv and don_gia_bh is not None:
                gia_duyet = gia_dvkt_dm.get(ma_dv)
                if gia_duyet is not None and abs(don_gia_bh - gia_duyet) > 0.5:
                    errors.append(_err(
                        ma_lk, "2. Danh mục - Giá - Định mức", "Lỗi",
                        "DON_GIA_BH", don_gia_bh,
                        f"Đơn giá DVKT '{ma_dv}' ({ten_dv}) = {don_gia_bh:,.0f} không khớp giá theo "
                        f"quy định hiện hành = {gia_duyet:,.0f}.",
                        "Sửa lại DON_GIA_BH theo đúng bảng giá dịch vụ kỹ thuật hiện hành (Thông tư giá).", nguon
                    ))

        if so_luong is not None and don_gia_bh is not None and thanh_tien_bh is not None:
            tinh_lai = round(so_luong * don_gia_bh, 0)
            if thanh_tien_bh - tinh_lai > max(1, 0.01 * tinh_lai):
                errors.append(_err(
                    ma_lk, "2. Danh mục - Giá - Định mức", "Lỗi",
                    "THANH_TIEN_BH", thanh_tien_bh,
                    f"Thành tiền BHYT ({thanh_tien_bh:,.0f}) lớn hơn Số lượng x Đơn giá BH "
                    f"({so_luong:g} x {don_gia_bh:,.0f} = {tinh_lai:,.0f}).",
                    "Kiểm tra lại công thức tính thành tiền BHYT của dịch vụ kỹ thuật.", nguon
                ))

        if so_luong is not None and so_luong <= 0:
            errors.append(_err(
                ma_lk, "2. Danh mục - Giá - Định mức", "Lỗi",
                "SO_LUONG", so_luong,
                "Số lượng DVKT phải lớn hơn 0.",
                "Kiểm tra lại số lượng dịch vụ kỹ thuật đã thực hiện.", nguon
            ))

        # 2.3 Cảnh báo định mức VTYT hao phí kèm DVKT (nếu có danh mục)
        dinh_muc = danh_muc.get("dinh_muc", {})
        if dinh_muc and ma_dv in dinh_muc:
            # Việc đối chiếu thực tế VTYT đã xuất kèm DVKT cần dữ liệu chi tiết
            # VTYT theo từng DVKT (thường nằm trong XML3 mở rộng / XML7).
            # Ở đây chỉ nhắc nhở để bác sĩ/điều dưỡng tự rà soát theo định mức.
            errors.append(_err(
                ma_lk, "2. Danh mục - Giá - Định mức", "Cảnh báo",
                "MA_DICH_VU", ma_dv,
                f"DVKT '{ma_dv}' ({ten_dv}) có định mức VTYT hao phí kèm theo - cần rà soát "
                "số lượng VTYT xuất kèm không vượt định mức quy định.",
                "Đối chiếu số lượng VTYT đã xuất cho DVKT này với định mức kinh tế kỹ thuật đã tải lên; "
                "nếu vượt định mức cần ghi rõ lý do (biến chứng, tai biến) trong hồ sơ bệnh án.", nguon
            ))

    return errors


# =====================================================================
#  NHÓM 3: HỢP LÝ CHỈ ĐỊNH Y KHOA (BẮT CHÉO DỮ LIỆU)
# =====================================================================

def check_nhom3_hop_ly_chi_dinh(ho_so_hc, file_type, items_by_malk, ma_lk, danh_muc=None):
    """
    items_by_malk: dict {ma_lk: {"thuoc": [...], "dvkt": [...], "cdha": [...]}}
                    - mỗi item là dict các trường gốc đã đọc được từ XML tương ứng,
                      dùng để kiểm tra trùng lặp / liều dùng / giới tính-tuổi.
    danh_muc có thể chứa:
      danh_muc['icd_chi_dinh'] : dict {MA_BENH (ICD10): set(MA_DICH_VU hoặc MA_THUOC cho phép)}
    Trả về list lỗi áp dụng cho TỪNG hồ sơ (ma_lk), được gọi 1 lần / ma_lk
    sau khi đã đọc xong tất cả các file.
    """
    errors = []
    danh_muc = danh_muc or {}
    if ho_so_hc is None:
        return errors

    ma_benh_chinh = _text(ho_so_hc, "MA_BENH_CHINH")
    ma_benh_kt = _text(ho_so_hc, "MA_BENH_KT")  # mã bệnh kèm theo (có thể nhiều mã, phân tách bằng ;)
    cac_ma_benh = set()
    if ma_benh_chinh:
        cac_ma_benh.add(ma_benh_chinh.upper())
    if ma_benh_kt:
        for m in re.split(r"[;,]", ma_benh_kt):
            m = m.strip().upper()
            if m:
                cac_ma_benh.add(m)

    gioi_tinh = _text(ho_so_hc, "GIOI_TINH")  # 1=Nam, 2=Nữ
    ngay_sinh = _text(ho_so_hc, "NGAY_SINH")
    ngay_vao = _text(ho_so_hc, "NGAY_VAO")

    tuoi = None
    dt_vao = _to_datetime(ngay_vao)
    if re.match(r"^\d{12}$", ngay_sinh or ""):
        dt_sinh = _to_datetime(ngay_sinh)
        if dt_sinh and dt_vao:
            tuoi = (dt_vao - dt_sinh).days // 365
    elif re.match(r"^\d{4}$", ngay_sinh or "") and dt_vao:
        try:
            tuoi = dt_vao.year - int(ngay_sinh)
        except ValueError:
            tuoi = None

    data = items_by_malk.get(ma_lk, {})

    # ---- 3.1 Chỉ định theo ICD-10 (bắt chéo với danh mục nếu có) ----
    icd_chi_dinh = danh_muc.get("icd_chi_dinh", {})
    for nhom_item, ds in (("dvkt", data.get("dvkt", [])), ("thuoc", data.get("thuoc", []))):
        for item in ds:
            ma_code = item.get("MA_DICH_VU") or item.get("MA_THUOC")
            ten_code = item.get("TEN_DICH_VU") or item.get("TEN_THUOC") or ""
            nguon = "XML3 - DVKT" if nhom_item == "dvkt" else "XML2 - Thuốc"
            truong_loi = "MA_DICH_VU" if nhom_item == "dvkt" else "MA_THUOC"

            if icd_chi_dinh:
                cho_phep = False
                for ma_benh in cac_ma_benh:
                    allowed_set = icd_chi_dinh.get(ma_benh) or icd_chi_dinh.get(ma_benh.split(".")[0])
                    if allowed_set and ma_code in allowed_set:
                        cho_phep = True
                        break
                if not cho_phep:
                    errors.append(_err(
                        ma_lk, "3. Hợp lý chỉ định y khoa", "Cảnh báo",
                        truong_loi, ma_code,
                        f"Chỉ định '{ma_code}' ({ten_code}) chưa thấy trong danh mục chỉ định phù hợp "
                        f"với mã bệnh chính/kèm theo ({', '.join(cac_ma_benh) or '(trống)'}).",
                        "Rà soát lại: nếu chỉ định này phù hợp với chẩn đoán, bổ sung mã bệnh kèm theo "
                        "(MA_BENH_KT) tương ứng vào hồ sơ hành chính; nếu không phù hợp, xem lại chỉ định "
                        "hoặc bổ sung lý do y khoa trong bệnh án để tránh bị xuất toán.", nguon
                    ))
            else:
                # Không có bảng đối chiếu ICD-DVKT: chỉ cảnh báo nhẹ cho các
                # DVKT chi phí cao thường bị soi, để bác sĩ tự rà soát ICD
                if nhom_item == "dvkt" and any(k in (ten_code or "").upper() for k in HIGH_COST_DVKT_KEYWORDS):
                    if not cac_ma_benh:
                        errors.append(_err(
                            ma_lk, "3. Hợp lý chỉ định y khoa", "Cảnh báo",
                            truong_loi, ma_code,
                            f"[Gợi ý theo từ khóa] Chỉ định '{ma_code}' ({ten_code}) là DVKT chi phí cao "
                            "nhưng hồ sơ chưa có mã bệnh chính (MA_BENH_CHINH) để đối chiếu mức độ hợp lý.",
                            "Bổ sung MA_BENH_CHINH (và MA_BENH_KT nếu có) trong XML1 - Hành chính.", nguon
                        ))

    # ---- 3.2 Chỉ định theo giới tính / tuổi ----
    for nhom_item, ds in (("dvkt", data.get("dvkt", [])), ("thuoc", data.get("thuoc", []))):
        for item in ds:
            ma_code = item.get("MA_DICH_VU") or item.get("MA_THUOC")
            ten_code = (item.get("TEN_DICH_VU") or item.get("TEN_THUOC") or "").upper()
            nguon = "XML3 - DVKT" if nhom_item == "dvkt" else "XML2 - Thuốc"
            truong_loi = "MA_DICH_VU" if nhom_item == "dvkt" else "MA_THUOC"

            if gioi_tinh == "1" and any(k in ten_code for k in FEMALE_ONLY_KEYWORDS):
                errors.append(_err(
                    ma_lk, "3. Hợp lý chỉ định y khoa", "Cảnh báo",
                    truong_loi, ma_code,
                    f"[Gợi ý theo từ khóa - cần xác nhận] Chỉ định '{ma_code}' ({ten_code.title()}) "
                    "có tên gợi ý dành cho bệnh nhân nữ, nhưng GIOI_TINH của hồ sơ là Nam (1).",
                    "Kiểm tra lại: nếu chỉ định này thực sự không phù hợp với giới tính người bệnh, "
                    "sửa lại chỉ định; nếu phù hợp (ví dụ tên dịch vụ trùng từ khóa nhưng bản chất khác), "
                    "có thể bỏ qua cảnh báo này.", nguon
                ))
            if gioi_tinh == "2" and any(k in ten_code for k in MALE_ONLY_KEYWORDS):
                errors.append(_err(
                    ma_lk, "3. Hợp lý chỉ định y khoa", "Cảnh báo",
                    truong_loi, ma_code,
                    f"[Gợi ý theo từ khóa - cần xác nhận] Chỉ định '{ma_code}' ({ten_code.title()}) "
                    "có tên gợi ý dành cho bệnh nhân nam, nhưng GIOI_TINH của hồ sơ là Nữ (2).",
                    "Kiểm tra lại: nếu chỉ định này thực sự không phù hợp với giới tính người bệnh, "
                    "sửa lại chỉ định; nếu phù hợp (ví dụ tên dịch vụ trùng từ khóa nhưng bản chất khác), "
                    "có thể bỏ qua cảnh báo này.", nguon
                ))
            if tuoi is not None and tuoi >= 16 and any(k in ten_code for k in PEDIATRIC_KEYWORDS):
                errors.append(_err(
                    ma_lk, "3. Hợp lý chỉ định y khoa", "Cảnh báo",
                    truong_loi, ma_code,
                    f"[Gợi ý theo từ khóa] Chỉ định '{ma_code}' ({ten_code.title()}) có dấu hiệu dành "
                    f"cho trẻ em, nhưng tuổi người bệnh ước tính ~{tuoi} tuổi (>=16).",
                    "Kiểm tra lại Ngày sinh và chỉ định/mã dịch vụ-thuốc cho phù hợp với độ tuổi người bệnh.", nguon
                ))

    # ---- 3.3 Trùng lặp chỉ định cùng loại trong ngày/đợt điều trị ----
    for nhom_item, label, key_field, name_field, nguon in (
        ("dvkt", "DVKT", "MA_DICH_VU", "TEN_DICH_VU", "XML3 - DVKT"),
        ("cdha", "CĐHA/TDCN", "MA_DICH_VU", "TEN_DICH_VU", "XML4 - CĐHA/TDCN"),
    ):
        ds = data.get(nhom_item, [])
        seen = {}
        for item in ds:
            ma_code = item.get(key_field)
            ten_code = item.get(name_field, "")
            ngay_yc = (item.get("NGAY_YL") or "")[:8]  # lấy phần ngày YYYYMMDD
            key = (ma_code, ngay_yc)
            if not ma_code:
                continue
            seen.setdefault(key, []).append(item)

        for (ma_code, ngay_yc), ds_dup in seen.items():
            if len(ds_dup) > 1:
                ten_code = ds_dup[0].get(name_field, "")
                errors.append(_err(
                    ma_lk, "3. Hợp lý chỉ định y khoa", "Cảnh báo",
                    key_field, ma_code,
                    f"Chỉ định '{ma_code}' ({ten_code}) - nhóm {label} được lặp lại "
                    f"{len(ds_dup)} lần trong cùng ngày {ngay_yc or '(không rõ)'}.",
                    "Nếu việc thực hiện lại là cần thiết về chuyên môn (ví dụ theo dõi diễn biến, "
                    "có biến chứng), cần ghi rõ lý do y khoa trong bệnh án để tránh bị xuất toán do trùng chỉ định.", nguon
                ))

    # ---- 3.4 Liều dùng và số ngày điều trị (thuốc) ----
    ngay_ra = _text(ho_so_hc, "NGAY_RA")
    dt_ra = _to_datetime(ngay_ra)
    so_ngay_dieu_tri = None
    if dt_vao and dt_ra:
        so_ngay_dieu_tri = max((dt_ra - dt_vao).days, 1)

    for item in data.get("thuoc", []):
        ma_thuoc = item.get("MA_THUOC")
        ten_thuoc = item.get("TEN_THUOC", "")
        lieu_dung = item.get("LIEU_DUNG", "")
        duong_dung = item.get("DUONG_DUNG", "")
        so_luong = _to_float(item.get("SO_LUONG"))

        if not lieu_dung:
            errors.append(_err(
                ma_lk, "3. Hợp lý chỉ định y khoa", "Cảnh báo",
                "LIEU_DUNG", "(thiếu)",
                f"Thuốc '{ma_thuoc}' ({ten_thuoc}) thiếu thông tin liều dùng (LIEU_DUNG).",
                "Bổ sung liều dùng/cách dùng (LIEU_DUNG, DUONG_DUNG) theo y lệnh để đối chiếu "
                "số lượng cấp phát với số ngày điều trị.", nguon
            ))
            continue

        # Tách số lần dùng/ngày từ chuỗi liều dùng dạng phổ biến, ví dụ:
        # "Uống ngày 2 lần, mỗi lần 1 viên" -> tổng = 2 viên/ngày
        so_lan = re.findall(r"(\d+(?:[.,]\d+)?)\s*(?:lần|l)\b", lieu_dung, flags=re.IGNORECASE)
        so_moi_lan = re.findall(r"mỗi lần\s*(\d+(?:[.,]\d+)?)", lieu_dung, flags=re.IGNORECASE)

        if so_lan and so_moi_lan and so_luong is not None and so_ngay_dieu_tri:
            try:
                tong_ngay = float(so_lan[0].replace(",", ".")) * float(so_moi_lan[0].replace(",", "."))
                du_kien = tong_ngay * so_ngay_dieu_tri
                # Cho phép sai lệch +/- 20% (dự trù mang theo khi xuất viện, đợt điều trị ngắt quãng...)
                if so_luong > du_kien * 1.2:
                    errors.append(_err(
                        ma_lk, "3. Hợp lý chỉ định y khoa", "Cảnh báo",
                        "SO_LUONG", so_luong,
                        f"Thuốc '{ma_thuoc}' ({ten_thuoc}): số lượng cấp = {so_luong:g}, theo liều dùng "
                        f"'{lieu_dung}' và {so_ngay_dieu_tri} ngày điều trị thì dự kiến chỉ cần "
                        f"~{du_kien:g}. Số lượng cấp đang vượt đáng kể so với dự kiến.",
                        "Kiểm tra lại số lượng thuốc kê/cấp so với liều dùng và số ngày điều trị "
                        "(NGAY_VAO - NGAY_RA); nếu có lý do hợp lý (mang về điều trị ngoại trú tiếp...) "
                        "cần ghi chú rõ trong bệnh án.", nguon
                    ))
            except (ValueError, ZeroDivisionError):
                pass

    return errors


# =====================================================================
#  NHÓM 4: THỜI GIAN KCB & TRÙNG LẶP ĐỢT ĐIỀU TRỊ
# =====================================================================

def check_nhom4_thoi_gian_trung_lap(ho_so, file_type, all_hanh_chinh=None):
    """
    all_hanh_chinh: list các dict thông tin hành chính đã đọc, mỗi dict gồm
                    ít nhất MA_LK, MA_THE_BHYT, NGAY_VAO, NGAY_RA, MA_CSKCB
                    -> dùng để phát hiện trùng lịch KCB cùng mã thẻ, các
                       khoảng thời gian giao nhau.
    Hàm này được gọi riêng cho XML1 - Hành chính, kiểm tra:
      4.1 Thời gian đẩy hồ sơ (NGAY_TT so với NGAY_RA)
      4.2 Trùng lịch KCB (chồng lấp thời gian, cùng mã thẻ, khác MA_LK)
      4.3 Tách đợt điều trị ngoại trú trong cùng 1 ngày tại cùng cơ sở
    """
    errors = []
    nguon = "XML1 - Hành chính"

    ma_lk_tag = ho_so.find('MA_LK')
    ma_lk = ma_lk_tag.text.strip().upper() if ma_lk_tag is not None and ma_lk_tag.text else "Không xác định"

    ngay_vao = _text(ho_so, "NGAY_VAO")
    ngay_ra = _text(ho_so, "NGAY_RA")
    ma_the = _text(ho_so, "MA_THE_BHYT")
    ma_cskcb = _text(ho_so, "MA_CSKCB")

    dt_vao = _to_datetime(ngay_vao)
    dt_ra = _to_datetime(ngay_ra)

    # ---- 4.1 Thời hạn đẩy hồ sơ ----
    ngay_tt = _text(ho_so, "NGAY_TT")  # ngày thanh toán/quyết toán (nếu có khai trong XML1)
    if ngay_tt and dt_ra:
        dt_tt = _to_datetime(ngay_tt)
        if dt_tt:
            so_ngay_cham = (dt_tt - dt_ra).days
            if so_ngay_cham > 10:
                errors.append(_err(
                    ma_lk, "4. Thời gian & trùng lặp KCB", "Cảnh báo",
                    "NGAY_TT", ngay_tt,
                    f"Khoảng cách từ ngày ra viện ({ngay_ra}) đến ngày quyết toán ({ngay_tt}) là "
                    f"{so_ngay_cham} ngày, có thể vượt thời hạn gửi dữ liệu lên Cổng giám định.",
                    "Đẩy hồ sơ lên Cổng tiếp nhận ngay sau khi người bệnh ra viện theo đúng thời hạn "
                    "quy định; nếu chậm cần kèm văn bản giải trình lý do.", nguon
                ))

    # ---- 4.2 & 4.3: cần dữ liệu tổng hợp tất cả hồ sơ ----
    if all_hanh_chinh and dt_vao and dt_ra and ma_the:
        for other in all_hanh_chinh:
            if other["ma_lk"] == ma_lk:
                continue
            if other["ma_the"] != ma_the:
                continue
            o_vao, o_ra = other["dt_vao"], other["dt_ra"]
            if not (o_vao and o_ra):
                continue

            # Kiểm tra giao nhau về thời gian
            giao_nhau = dt_vao <= o_ra and o_vao <= dt_ra
            if giao_nhau:
                if other["ma_cskcb"] != ma_cskcb:
                    errors.append(_err(
                        ma_lk, "4. Thời gian & trùng lặp KCB", "Lỗi",
                        "NGAY_VAO/NGAY_RA", f"{ngay_vao} - {ngay_ra}",
                        f"Khoảng thời gian KCB ({ngay_vao} - {ngay_ra}) trùng/giao với một hồ sơ khác "
                        f"(MA_LK={other['ma_lk']}) tại cơ sở khác (MA_CSKCB={other['ma_cskcb']}) "
                        f"cùng mã thẻ BHYT {ma_the}.",
                        "Kiểm tra lại thời gian vào/ra viện; nếu là khám chuyên khoa đặc thù được phép "
                        "song song (chạy thận, hoá trị...) cần lưu hồ sơ chứng minh; nếu không, "
                        "phối hợp với cơ sở liên quan để xác minh và điều chỉnh.", nguon
                    ))
                else:
                    # Cùng 1 cơ sở, cùng ngày, nhiều đợt khám -> kiểm tra tách đợt ngoại trú
                    if dt_vao.date() == o_vao.date() and ngay_vao != other["ngay_vao"]:
                        errors.append(_err(
                            ma_lk, "4. Thời gian & trùng lặp KCB", "Cảnh báo",
                            "NGAY_VAO", ngay_vao,
                            f"Cùng ngày {dt_vao.date()} tại cùng cơ sở (MA_CSKCB={ma_cskcb}), người bệnh "
                            f"có nhiều hồ sơ KCB (MA_LK={ma_lk} và {other['ma_lk']}).",
                            "Khám nhiều chuyên khoa trong cùng một ngày tại một cơ sở chỉ được tính là "
                            "một đợt khám (trừ trường hợp được phép tách riêng theo quy định). "
                            "Kiểm tra gộp thành 1 đợt khám nếu không thuộc trường hợp ngoại lệ.", nguon
                        ))

    return errors


# =====================================================================
#  HÀM TỔNG HỢP: CHẠY TẤT CẢ QUY TẮC CHO 1 HỒ SƠ HÀNH CHÍNH
#  (dùng cho nhóm 1 và 4, gọi từ vòng lặp chính trong app)
# =====================================================================

def check_hanh_chinh_full(ho_so, all_hanh_chinh=None, danh_muc=None):
    ma_lk_tag = ho_so.find('MA_LK')
    ma_lk = ma_lk_tag.text.strip().upper() if ma_lk_tag is not None and ma_lk_tag.text else "Không xác định"

    errors = []
    errors.extend(check_nhom1_hanh_chinh(ho_so, ma_lk, danh_muc=danh_muc))
    errors.extend(check_nhom4_thoi_gian_trung_lap(ho_so, "Hành chính", all_hanh_chinh=all_hanh_chinh))
    return errors
