"""
Công cụ sửa lỗi & kiểm soát quy tắc dữ liệu KSK
================================================
Áp dụng cho 3 mẫu: Dưới 6 tuổi, 6 đến dưới 18 tuổi, Trên 18 tuổi.

CÁC LỖI ĐÃ SỬA (đánh số để dễ đối chiếu khi review):

[FIX 1] Named Range không được nhận diện làm droplist.
    Bản gốc chỉ xử lý formula1 dạng '"A,B,C"' hoặc 'Sheet!Range'. Nhưng gần hết
    droplist quan trọng của Master (GioiTinh, DM_dantoc, DM_DTkham, TINH,
    DM_Icd10...) là Named Range thuần (formula1 chỉ ghi tên, không có dấu !).
    -> Thêm resolve_named_range_values() để tra cứu qua wb.defined_names.

[FIX 2] Cột không có tên biến (row 2 = None) bị gộp chung và ghi đè lẫn nhau.
    Các cột công thức ẩn (tra tên tỉnh/tên xã tự động bằng INDEX/MATCH) và cột
    STT đều có var_name = None. Dùng None làm key dict khiến chúng bị gộp làm
    một, và khi xuất file sẽ bị ghi đè bằng cùng 1 giá trị -> XÓA MẤT CÔNG THỨC
    tra tên tỉnh/xã trong file xuất ra.
    -> Bỏ qua hoàn toàn các cột này (không đọc, không ghi), giữ nguyên như
    trong Master gốc.

[FIX 3] Tên biến trùng lặp trong chính Master (vd DIEN_THOAI xuất hiện 2 lần
    ở file 6-18 tuổi, cột I và cột W) làm dữ liệu ghi đè lẫn nhau.
    -> Dùng col_idx làm khóa nội bộ (internal_key), chỉ thêm hậu tố khi thật
    sự trùng tên, và hiển thị rõ số cột trong nhãn để người dùng phân biệt.

[FIX 4] Rule kiểm tra CCCD áp cho tên biến sai (SO_DINH_DANH/MA_SO_BHXH...)
    trong khi field thật là SO_CCCD, SO_CCCD_NGH, SO_CCCD_NGUOI_DI_CUNG.
    Đồng thời tránh khớp nhầm NGAYCAP_CCCD (ngày cấp)/NOICAP_CCCD (nơi cấp).
    -> Chỉ khớp field bắt đầu bằng SO_/MA_SO_ và có chứa CCCD/DINH_DANH/BHXH.

[FIX 5] Rule Giới tính hard-code chỉ nhận "Nam"/"Nữ", trong khi droplist thật
    của Master có thêm "Chưa xác định".
    -> Ưu tiên validate theo đúng droplist đọc được từ Master; hard-code chỉ
    dùng cho các field không có droplist (CCCD, ngày tháng).

[FIX 6] Không cảnh báo khi file lỗi thiếu hẳn 1 cột so với Master.
    -> Chặn ngay từ bước tải file, liệt kê rõ cột nào bị thiếu.

[FIX 7] Không cảnh báo khi số dòng dữ liệu vượt quá số dòng đã được định dạng
    sẵn (có công thức/droplist) trong Master -> dòng dư có thể thiếu công thức
    tra tên tỉnh/xã.
    -> Cảnh báo rõ ràng trước khi tải file kết quả.

[Vấn đề 1] NGAY_SINH bị tô màu lỗi nhưng ngày vẫn đúng định dạng/hợp lệ có thể
    là do chọn nhầm mẫu tuổi (AGE_TEMPLATE_MISMATCH) chứ không phải lỗi nhập
    liệu -> không bắt buộc sửa, chỉ nhắc nhở, và tô lavender cả dòng khi export.

[Mẫu Trên 18 tuổi] Thêm nhánh loại mẫu thứ 3 (nhận diện qua cột đặc trưng
    CAC_BENH_TAT_NEU_CO), tái sử dụng cơ chế "note column" có sẵn.

[Rà soát ổn định] MA_NGHE_NGHIEP và MAXA_CU_TRU không load được droplist:
    - MA_NGHE_NGHIEP dùng Named Range dạng công thức OFFSET/COUNTIF động
      (openpyxl không tính được kết quả công thức) -> đọc trực tiếp dữ liệu
      từ sheet danh mục dựa theo ô neo trong OFFSET().
    - MAXA_CU_TRU dùng droplist "cascading" (INDIRECT+SUBSTITUTE trỏ theo
      giá trị MATINH_CU_TRU của từng dòng, mỗi tỉnh có 1 Named Range xã
      riêng) -> gộp toàn bộ xã của tất cả các tỉnh thành 1 danh sách chung.
    - format_cell_value xử lý an toàn ô Excel dạng Number (bỏ đuôi ".0").
    - Bọc try/except quanh việc mở file Master (cả lúc xử lý file lỗi và
      lúc xuất file) -> không crash toàn app nếu Master hỏng/bị khoá; ở
      bước export, dữ liệu người dùng đã sửa KHÔNG bị mất nếu export lỗi.
    - Dọn sạch widget key cũ khi tải file mới, tránh dữ liệu dính giữa
      các lần tải file liên tiếp.

[Đa giá trị ICD-10] KET_LUAN_BENH, TSGD_MA_BENH, MA_BENH_SAN_KHOA_KHONG_BT,
    TSBT_MA_BENH là field ĐA GIÁ TRỊ (nhiều mã ICD-10 cách nhau bằng ';').
    validate_field_value bỏ qua rule so khớp droplist nguyên chuỗi cho các
    field này (nếu không sẽ luôn báo sai với giá trị có từ 2 mã trở lên).
    Giao diện dùng ĐÚNG 1 ô nhập liệu duy nhất (st.text_input) làm nguồn dữ
    liệu thật, cho gõ tay tự do — không tách thành nhiều widget/nhiều dòng.
    Kèm 1 ô tra cứu (st.selectbox, gõ để tìm kiếm/autocomplete, tận dụng lại
    droplist ICD-10 đã tải sẵn từ Master) + nút "➕ Thêm": bấm nút mới trích
    riêng phần MÃ (trước dấu " - ") và nối vào ô nhập liệu chính bằng ';',
    không trùng mã đã có, không dư ';' đầu/cuối, rồi tự làm mới ô tra cứu.
    Việc ghi giá trị + reset CHỈ xảy ra khi bấm nút (không tự động theo lựa
    chọn của selectbox) — dùng đúng pattern nút+rerun đã ổn định nhiều lần
    trong app này, tránh lặp lại sự cố rerun tự phát đã từng xảy ra ở phiên
    bản multiselect trước đó.
"""

from collections import defaultdict
from datetime import datetime
import io
import os
import re

import openpyxl  # type: ignore
from openpyxl.styles import PatternFill  # type: ignore
from openpyxl.utils import get_column_letter, column_index_from_string  # type: ignore
import pandas as pd  # type: ignore
import streamlit as st  # type: ignore

# ==========================================================================
# CẤU HÌNH TRANG & STYLE
# ==========================================================================
st.set_page_config(
    page_title="MiniTool_Chinh_Loi_Ksk_Chuan_MasterFlow", layout="wide"
)

st.markdown(
    """
    <style>
    .main-title { font-size: 24px; font-weight: bold; color: #1E3A8A; margin-bottom: 10px; }
    .stButton>button { width: 100%; border-radius: 6px; }
    .rule-warning {
        padding: 8px; background-color: #FEF3C7; border-left: 4px solid #F59E0B;
        margin-bottom: 10px; border-radius: 4px; color: #92400E;
    }
    .error-block {
        padding: 10px; background-color: #FEE2E2; border-left: 4px solid #EF4444;
        margin-bottom: 10px; border-radius: 4px; color: #B91C1C;
    }
    .note-block {
        padding: 10px; background-color: #EFF6FF; border-left: 4px solid #3B82F6;
        margin-bottom: 10px; border-radius: 4px; color: #1E40AF; font-size: 13px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


def reset_app_state():
    st.session_state.uploader_key += 1
    keys_to_delete = [
        "df_data", "err_reg", "initial_err_reg", "m_type", "master_fields",
        "field_meta", "note_col_key", "file_name", "status_completed",
        "export_file_ready", "person_selector_radio", "person_selector_target",
        "rule_warnings", "action_error_msg", "dropdown_options",
        "dropdown_lookup_sets", "original_input_values", "row_capacity_warning",
        "age_mismatch_rows", "template_data_rows", "master_path",
    ]
    for k in list(st.session_state.keys()):
        if k in keys_to_delete or k.startswith("key_"):
            del st.session_state[k]
    st.rerun()


col_header_1, col_header_2 = st.columns([4, 1])
with col_header_1:
    st.markdown(
        "<div class='main-title'>🛠️ SỬA THEO QUY TẮC FILE IMPORT DỮ LIỆU KSK CỔNG MASTERFLOW DONGNAI</div>",
        unsafe_allow_html=True,
    )
with col_header_2:
    if st.button("🔄 Tải file mới / Làm mới", type="secondary"):
        reset_app_state()


def find_master_file(base_name):
    for ext in [".xlsm", ".xlsx"]:
        path = base_name + ext
        if os.path.exists(path):
            return path
    return None


PATH_MASTER_UNDER6 = find_master_file("Import_KSK_Duoi 6 - MasterFlow")
PATH_MASTER_6TO18 = find_master_file("Import_KSK_6 den Duoi 18 - MasterFlow")
PATH_MASTER_TREN18 = find_master_file("Import_KSK_Tren18_MasterFlow")


# ==========================================================================
# CÁC HÀM DÙNG CHUNG
# ==========================================================================

def build_master_fields(ws):
    """Duyệt toàn bộ cột của sheet, trả về danh sách field hợp lệ.

    [FIX 2] Cột nào không có tên biến ở row 2 (var_name is None/rỗng) sẽ bị
    BỎ QUA hoàn toàn — đây thường là cột công thức ẩn (tra tên tỉnh/xã tự
    động) hoặc cột STT không cần đọc/ghi.

    [FIX 3] Nếu tên biến bị trùng trong cùng 1 sheet (vd DIEN_THOAI x2), mỗi
    lần xuất hiện sẽ có internal_key riêng (var_name, var_name__1, ...) để
    không bị ghi đè lẫn nhau, đồng thời label hiển thị kèm số cột để người
    dùng phân biệt được.
    """
    fields = []
    name_seen = {}
    for c in range(1, ws.max_column + 1):
        raw_var = ws.cell(row=2, column=c).value
        if raw_var is None or str(raw_var).strip() == "":
            continue
        var_name = str(raw_var).strip()
        display_name = ws.cell(row=1, column=c).value
        occurrence = name_seen.get(var_name, 0)
        name_seen[var_name] = occurrence + 1
        internal_key = var_name if occurrence == 0 else f"{var_name}__{occurrence}"
        col_letter = get_column_letter(c)
        label = var_name if occurrence == 0 else f"{var_name} (cột {col_letter})"
        fields.append(
            {
                "col_idx": c,
                "col_letter": col_letter,
                "var_name": var_name,
                "display_name": display_name,
                "internal_key": internal_key,
                "label": label,
            }
        )
    return fields


def resolve_named_range_values(wb, name):
    """[FIX 1] Đọc giá trị của 1 Excel Named Range (Defined Name) theo tên,
    dùng cho các validation formula1 chỉ ghi tên vùng (vd 'GioiTinh',
    'DM_dantoc') thay vì 'Sheet!$A$1:$A$5'."""
    try:
        dn = wb.defined_names[name]
    except Exception:
        return []
    values = []
    try:
        destinations = list(dn.destinations)
    except Exception:
        destinations = []

    if not destinations:
        # [Vấn đề 3] Named Range dùng công thức vùng động kiểu
        # OFFSET(Sheet!$A$2,0,0,COUNTIF(Sheet!$A:$A,"<>")-1,1) (vd MA_NGHE_NGHIEP
        # -> Named Range "NgheNghiep") sẽ có destinations rỗng vì openpyxl không
        # tính được kết quả OFFSET/COUNTIF. Suy ra vùng dữ liệu thực tế bằng cách
        # đọc toàn bộ giá trị không rỗng của đúng cột đó, từ ô neo trong OFFSET().
        formula_text = str(getattr(dn, "attr_text", "") or getattr(dn, "value", "") or "")
        m = re.match(r"OFFSET\(\s*([^!]+)!\$?([A-Z]+)\$?(\d+)", formula_text, re.IGNORECASE)
        if m:
            sheet_title = m.group(1).strip("'").strip('"').strip()
            if sheet_title in wb.sheetnames:
                ws_ref = wb[sheet_title]
                col_idx = column_index_from_string(m.group(2))
                start_row = int(m.group(3))
                for r in range(start_row, ws_ref.max_row + 1):
                    v = ws_ref.cell(row=r, column=col_idx).value
                    if v is not None and str(v).strip():
                        values.append(str(v).strip())
        return values

    for sheet_title, coord in destinations:
        if sheet_title not in wb.sheetnames:
            continue
        ws_ref = wb[sheet_title]
        coord_clean = coord.replace("$", "")
        try:
            cells = ws_ref[coord_clean]
        except Exception:
            continue
        rows_iter = cells if isinstance(cells, tuple) else [cells]
        for row in rows_iter:
            row_cells = row if isinstance(row, tuple) else [row]
            for cell in row_cells:
                if cell.value is not None:
                    values.append(str(cell.value).strip())
    return values


def format_cell_value(val):
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    if isinstance(val, float):
        # [Rà soát] Ô Excel bị gõ dạng Number thay vì Text (hay gặp với
        # CCCD/mã số) sẽ được openpyxl đọc thành float -> str() cho ra đuôi
        # ".0" gây sai lệch validate CCCD/ngày tháng. Bỏ phần thập phân nếu
        # là số nguyên.
        if val.is_integer():
            return str(int(val))
        return str(val).strip()
    val_str = str(val).strip()
    if val_str.lower() == "nan":
        return ""
    return val_str


def is_error_fill(fill):
    """Coi BẤT KỲ ô nào có tô màu khác trắng là ô lỗi (không phụ thuộc đúng
    1 mã hồng cụ thể, vì phần mềm quét có thể dùng nhiều sắc khác nhau)."""
    if fill is None:
        return False
    if getattr(fill, "fill_type", None) != "solid":
        return False
    color = fill.start_color
    if color is None or not color.rgb or not isinstance(color.rgb, str):
        return False
    rgb = color.rgb.upper()
    if len(rgb) == 8:  # ARGB -> RGB
        rgb = rgb[2:]
    return rgb != "FFFFFF"


CCCD_PATTERN = re.compile(r"^(SO_|MA_SO_).*(CCCD|DINH_DANH|BHXH)", re.IGNORECASE)
DATE_PATTERN = re.compile(r"^(0[1-9]|[12][0-9]|3[01])/(0[1-9]|1[0-2])/\d{4}$")

# [Đa giá trị ICD-10] Các field dùng chung droplist DM_Icd10 nhưng bản chất là
# ĐA GIÁ TRỊ — nhiều mã ICD-10 cách nhau bằng dấu ';' (vd "H52.1;J02.0;I10"),
# không phải droplist chọn 1 giá trị đơn. Cho phép gõ tay tự do hoặc chọn kết
# hợp nhiều mã từ danh mục (xem nhánh riêng trong validate_field_value và ở
# phần render giao diện bên dưới).
MULTI_VALUE_ICD10_FIELDS = {
    "KET_LUAN_BENH",
    "TSGD_MA_BENH",
    "MA_BENH_SAN_KHOA_KHONG_BT",
    "TSBT_MA_BENH",
}


def validate_field_value(var_name, val_str, dropdown_options, dropdown_lookup_sets):
    """[FIX 4][FIX 5] Ưu tiên validate theo đúng droplist thật đọc từ Master.
    Chỉ dùng rule hard-code (CCCD / ngày tháng) cho field không có droplist."""
    if not val_str:
        return None

    if var_name in MULTI_VALUE_ICD10_FIELDS:
        # [Đa giá trị ICD-10] Không áp rule so khớp NGUYÊN CHUỖI với 1 mã
        # ICD-10 (sẽ luôn sai với giá trị có từ 2 mã trở lên). Giao diện chọn
        # từ danh mục ICD-10 nằm ở phần render bên dưới.
        return None

    if var_name in dropdown_lookup_sets and dropdown_lookup_sets[var_name]:
        if val_str not in dropdown_lookup_sets[var_name]:
            opts = dropdown_options.get(var_name, [])
            if len(opts) <= 8:
                gợi_ý = ", ".join(opts)
                return f"Giá trị '{val_str}' không có trong danh mục chuẩn của Master. Chỉ nhận: {gợi_ý}."
            return f"Giá trị '{val_str}' không có trong danh mục chuẩn của Master (cột {var_name})."
        return None

    if CCCD_PATTERN.search(var_name):
        if not re.match(r"^\d{10}$", val_str) and not re.match(r"^\d{12}$", val_str):
            return f"Số CCCD/Định danh/BHXH phải gồm 10 hoặc 12 chữ số (hiện tại: '{val_str}')."
        return None

    if "NGAY" in var_name:
        if not DATE_PATTERN.match(val_str):
            return f"Ngày tháng phải đúng định dạng DD/MM/YYYY có dấu '/' (hiện tại: '{val_str}')."
        return None

    return None


# ==========================================================================
# ĐỌC CẤU TRÚC MASTER: field list + droplist chuẩn
# ==========================================================================
@st.cache_data
def load_master_structure(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True, keep_vba=True)
    sheet_name = [
        s for s in wb.sheetnames
        if s not in ["DM_CanLamSang", "Hướng dẫn"] and not s.startswith("dm")
    ][0]
    ws = wb[sheet_name]

    master_fields = build_master_fields(ws)
    col_to_field = {f["col_idx"]: f for f in master_fields}

    dropdown_options = {}
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            formula = str(dv.formula1 or "").strip()
            if not formula:
                continue
            clean_formula = formula.lstrip("=").strip()
            options = []

            if clean_formula.startswith('"') and clean_formula.endswith('"'):
                options = [
                    opt.strip() for opt in clean_formula.strip('"').split(",") if opt.strip()
                ]
            elif "!" in clean_formula:
                try:
                    sheet_ref, range_ref = clean_formula.split("!", 1)
                    sheet_ref = sheet_ref.strip("'").strip('"').strip()
                    if sheet_ref in wb.sheetnames:
                        ws_ref = wb[sheet_ref]
                        range_clean = range_ref.replace("$", "")
                        cells = ws_ref[range_clean]
                        rows_iter = cells if isinstance(cells, tuple) else [cells]
                        for row in rows_iter:
                            row_cells = row if isinstance(row, tuple) else [row]
                            for cell in row_cells:
                                if cell.value is not None:
                                    options.append(str(cell.value).strip())
                except Exception as e:
                    print(f"Lỗi đọc droplist liên sheet cho công thức {formula}: {e}")
            else:
                # [FIX 1] Named Range thuần (không dấu !, không dấu ")
                options = resolve_named_range_values(wb, clean_formula)

            if not options:
                continue

            unique_opts = []
            for opt in options:
                if opt and opt not in unique_opts:
                    unique_opts.append(opt)

            sqref_str = str(dv.sqref)
            for part in sqref_str.split():
                match = re.match(r"([A-Z]+)", part)
                if not match:
                    continue
                col_idx = column_index_from_string(match.group(1))
                field = col_to_field.get(col_idx)
                if field:
                    dropdown_options[field["var_name"]] = unique_opts

    # [Vấn đề 3] MAXA_CU_TRU dùng droplist "cascading" (INDIRECT + SUBSTITUTE
    # trỏ theo giá trị MATINH_CU_TRU của từng dòng) nên vòng lặp data_validations
    # ở trên không tự suy ra được 1 danh sách cố định (mỗi tỉnh có 1 Named Range
    # xã riêng, ví dụ tỉnh "Hà Nội" -> Named Range "Hà_Nội"). Gộp toàn bộ xã của
    # tất cả các tỉnh trong droplist MATINH_CU_TRU thành 1 danh sách chung để
    # hiển thị đầy đủ trên giao diện (đánh đổi: không lọc riêng theo từng tỉnh).
    if "MAXA_CU_TRU" not in dropdown_options and "MATINH_CU_TRU" in dropdown_options:
        xa_all = []
        xa_seen = set()
        for tinh in dropdown_options["MATINH_CU_TRU"]:
            ten_named_range = tinh.replace(" ", "_").replace(".", "_")
            for xa in resolve_named_range_values(wb, ten_named_range):
                if xa not in xa_seen:
                    xa_seen.add(xa)
                    xa_all.append(xa)
        if xa_all:
            dropdown_options["MAXA_CU_TRU"] = xa_all

    dropdown_lookup_sets = {k: set(v) for k, v in dropdown_options.items()}
    template_data_rows = max(0, ws.max_row - 3)  # dữ liệu bắt đầu từ dòng 4

    return {
        "sheet_name": sheet_name,
        "master_fields": master_fields,
        "dropdown_options": dropdown_options,
        "dropdown_lookup_sets": dropdown_lookup_sets,
        "template_data_rows": template_data_rows,
    }


# ==========================================================================
# XỬ LÝ FILE LỖI NGƯỜI DÙNG TẢI LÊN
# ==========================================================================

def _ngay_sinh_age_template_mismatch(val_str, master_type):
    """[Vấn đề 1] Dùng riêng cho NGAY_SINH khi giá trị đã đúng định dạng
    dd/mm/yyyy (rule_err is None): kiểm tra tuổi tính từ NGAY_SINH có phù hợp
    với Master đang chọn (UNDER6 / 6TO18) hay không -> AGE_TEMPLATE_MISMATCH.
    Không thay đổi/ảnh hưởng đến validate_field_value hiện tại."""
    try:
        day, month, year = (int(p) for p in val_str.split("/"))
        birth = datetime(year, month, day)
    except (ValueError, TypeError):
        return False
    today = datetime.now()
    age = today.year - birth.year - ((today.month, today.day) < (birth.month, birth.day))
    if master_type == "UNDER6":
        return age >= 6
    if master_type == "6TO18":
        return age < 6 or age >= 18
    return False


def process_uploaded_error_file(uploaded_file):
    try:
        wb_err = openpyxl.load_workbook(uploaded_file, data_only=True)
        ws_err = wb_err.active
    except Exception as e:
        return None, f"LỖI: Không thể đọc file Excel. Chi tiết: {e}"

    err_fields = build_master_fields(ws_err)
    err_var_names = {f["var_name"] for f in err_fields}

    if "GHI_RO_VAN_DE_SUC_KHOE" in err_var_names:
        master_type = "UNDER6"
        master_path = PATH_MASTER_UNDER6
        note_col_var = "GHI_RO_VAN_DE_SUC_KHOE"
    elif "CAC_VAN_DE_SUC_KHOE" in err_var_names:
        master_type = "6TO18"
        master_path = PATH_MASTER_6TO18
        note_col_var = "CAC_VAN_DE_SUC_KHOE"
    elif "CAC_BENH_TAT_NEU_CO" in err_var_names:
        master_type = "TREN18"
        master_path = PATH_MASTER_TREN18
        note_col_var = "CAC_BENH_TAT_NEU_CO"
    else:
        return None, "CẢNH BÁO: File không đúng định dạng mẫu KSK chuẩn!"

    if not master_path or not os.path.exists(master_path):
        target_name = (
            "Import_KSK_Duoi 6 - MasterFlow" if master_type == "UNDER6"
            else "Import_KSK_6 den Duoi 18 - MasterFlow" if master_type == "6TO18"
            else "Import_KSK_Tren18_MasterFlow"
        )
        return None, f"LỖI: Không tìm thấy file mẫu '{target_name}' trong thư mục!"

    try:
        master = load_master_structure(master_path)
    except Exception as e:
        return None, (
            f"LỖI: Không thể đọc file mẫu Master '{os.path.basename(master_path)}'. "
            f"File có thể bị hỏng, sai định dạng, hoặc đang được mở bởi chương trình khác. "
            f"Chi tiết: {e}"
        )
    master_fields = master["master_fields"]
    dropdown_options = master["dropdown_options"]
    dropdown_lookup_sets = master["dropdown_lookup_sets"]

    # [FIX 3] Khớp cột theo tên biến, xử lý cả trường hợp tên biến trùng lặp
    # bằng cách khớp theo đúng thứ tự xuất hiện (occurrence).
    err_cols_by_name = defaultdict(list)
    for f in err_fields:
        err_cols_by_name[f["var_name"]].append(f["col_idx"])

    name_seen = {}
    resolved = []  # list of (master_field, err_col_idx_or_None)
    for mf in master_fields:
        occ = name_seen.get(mf["var_name"], 0)
        name_seen[mf["var_name"]] = occ + 1
        candidates = err_cols_by_name.get(mf["var_name"], [])
        err_col = candidates[occ] if occ < len(candidates) else None
        resolved.append((mf, err_col))

    # [FIX 6] Chặn sớm nếu file lỗi thiếu cột bắt buộc so với Master
    missing = [mf["label"] for mf, err_col in resolved if err_col is None]
    if missing:
        shown = missing[:15]
        extra = f" (và {len(missing)-15} cột khác)" if len(missing) > 15 else ""
        return None, (
            "CẢNH BÁO: File không đúng mẫu! Thiếu các cột bắt buộc: "
            + ", ".join(shown) + extra
        )

    ho_ten_err_col = None
    for f in err_fields:
        if f["var_name"] == "HO_TEN":
            ho_ten_err_col = f["col_idx"]
            break

    data_list = []
    error_registry = {}
    rule_warnings = {}
    age_mismatch_rows = set()
    field_meta = {mf["internal_key"]: mf for mf, _ in resolved}

    for r in range(4, ws_err.max_row + 1):
        ho_ten_str = ""
        if ho_ten_err_col:
            ho_ten_str = format_cell_value(ws_err.cell(row=r, column=ho_ten_err_col).value)
        if not ho_ten_str:
            continue

        row_data = {}
        row_errors = []
        row_rule_warns = {}
        row_age_mismatch = False

        for mf, err_col in resolved:
            cell = ws_err.cell(row=r, column=err_col)
            val_str = format_cell_value(cell.value)
            row_data[mf["internal_key"]] = val_str

            is_note_col = mf["var_name"] == note_col_var
            has_error_fill = is_error_fill(cell.fill)

            rule_err = validate_field_value(
                mf["var_name"], val_str, dropdown_options, dropdown_lookup_sets
            )
            if rule_err and not is_note_col:
                row_rule_warns[mf["internal_key"]] = rule_err

            # [Vấn đề 1] NGAY_SINH bị tô màu (lỗi) nhưng ngày vẫn đúng định
            # dạng/hợp lệ (rule_err is None) -> có thể là lệch mẫu tuổi
            # (AGE_TEMPLATE_MISMATCH), không bắt buộc sửa lại ngày sinh.
            age_mismatch = False
            if mf["var_name"] == "NGAY_SINH" and has_error_fill and not rule_err:
                age_mismatch = _ngay_sinh_age_template_mismatch(val_str, master_type)
                if age_mismatch:
                    row_age_mismatch = True

            if not is_note_col and (has_error_fill or rule_err) and not age_mismatch:
                row_errors.append(mf["internal_key"])

        data_list.append(row_data)
        r_idx = len(data_list) - 1
        if row_errors:
            error_registry[r_idx] = list(dict.fromkeys(row_errors))
        if row_rule_warns:
            rule_warnings[r_idx] = row_rule_warns
        if row_age_mismatch:
            age_mismatch_rows.add(r_idx)

    df_data = pd.DataFrame(data_list)
    initial_error_registry = {k: list(v) for k, v in error_registry.items()}

    return {
        "master_type": master_type,
        "df_data": df_data,
        "err_reg": error_registry,
        "initial_err_reg": initial_error_registry,
        "rule_warnings": rule_warnings,
        "master_fields": master_fields,
        "field_meta": field_meta,
        "dropdown_options": dropdown_options,
        "dropdown_lookup_sets": dropdown_lookup_sets,
        "note_col_key": note_col_var,
        "template_data_rows": master["template_data_rows"],
        "master_path": master_path,
        "age_mismatch_rows": age_mismatch_rows,
    }, None


# ==========================================================================
# UI: TẢI FILE
# ==========================================================================
uploaded_file = st.file_uploader(
    "📂 Kéo thả hoặc chọn File Excel báo lỗi (.xlsx, .xlsm)",
    type=["xlsx", "xlsm"],
    key=f"file_uploader_{st.session_state.uploader_key}",
)

if uploaded_file is None and "df_data" in st.session_state:
    reset_app_state()

if uploaded_file is not None:
    if "file_name" not in st.session_state or st.session_state.file_name != uploaded_file.name:
        with st.spinner("Hệ thống đang quét ô báo lỗi và chuẩn hóa file theo Master Flow..."):
            result, err_msg = process_uploaded_error_file(uploaded_file)
            if err_msg:
                st.error(err_msg)
            else:
                # [Rà soát] Dọn sạch widget key (key_input_*/key_note_*) của file
                # trước đó trước khi nạp file mới — tránh giá trị cũ dính sang
                # file mới nếu trùng chỉ số dòng/tên field giữa 2 lần tải file
                # liên tiếp mà không bấm "Làm mới".
                for k in list(st.session_state.keys()):
                    if k.startswith("key_"):
                        del st.session_state[k]

                st.session_state.m_type = result["master_type"]
                st.session_state.df_data = result["df_data"]
                st.session_state.err_reg = result["err_reg"]
                st.session_state.initial_err_reg = result["initial_err_reg"]
                st.session_state.rule_warnings = result["rule_warnings"]
                st.session_state.master_fields = result["master_fields"]
                st.session_state.field_meta = result["field_meta"]
                st.session_state.dropdown_options = result["dropdown_options"]
                st.session_state.dropdown_lookup_sets = result["dropdown_lookup_sets"]
                st.session_state.note_col_key = result["note_col_key"]
                st.session_state.template_data_rows = result["template_data_rows"]
                st.session_state.master_path = result["master_path"]
                st.session_state.age_mismatch_rows = result["age_mismatch_rows"]
                st.session_state.file_name = uploaded_file.name
                st.session_state.status_completed = set()
                st.session_state.action_error_msg = ""

                st.session_state.original_input_values = {}
                for idx, row in result["df_data"].iterrows():
                    st.session_state.original_input_values[idx] = row.to_dict()

                n_rows = len(result["df_data"])
                if n_rows > result["template_data_rows"]:
                    st.session_state.row_capacity_warning = (
                        f"⚠️ File có {n_rows} dòng dữ liệu, vượt quá {result['template_data_rows']} "
                        "dòng đã được định dạng sẵn (công thức tra tên tỉnh/xã, droplist) trong Master. "
                        "Các dòng vượt quá có thể thiếu công thức/droplist khi xuất file — nên mở rộng "
                        "định dạng của Master xuống thêm dòng trước khi dùng, hoặc kiểm tra kỹ các dòng cuối "
                        "sau khi tải file kết quả về."
                    )
                else:
                    st.session_state.row_capacity_warning = None
                st.rerun()


def save_and_revalidate_person(current_idx, fields_to_save, note_key):
    init_errs_for_person = st.session_state.get("initial_err_reg", {}).get(current_idx, [])
    dropdown_options = st.session_state.dropdown_options
    dropdown_lookup_sets = st.session_state.dropdown_lookup_sets
    field_meta = st.session_state.field_meta

    remaining_errors = []
    new_rule_warns = {}

    for internal_key in fields_to_save:
        var_name = field_meta[internal_key]["var_name"]
        key_input = f"key_input_{current_idx}_{internal_key}"
        if key_input in st.session_state:
            new_val = str(st.session_state[key_input]).strip()
            st.session_state.df_data.at[current_idx, internal_key] = new_val

            if internal_key in init_errs_for_person:
                orig_val = str(
                    st.session_state.original_input_values.get(current_idx, {}).get(internal_key, "")
                ).strip()
                if new_val == orig_val or not new_val:
                    new_rule_warns[internal_key] = (
                        "Trường này đang bị lỗi/tô màu từ file gốc. Bạn bắt buộc phải nhập/sửa lại dữ liệu hợp lệ!"
                    )
                    if internal_key not in remaining_errors:
                        remaining_errors.append(internal_key)

        val_str = str(st.session_state.df_data.at[current_idx, internal_key]).strip()
        rule_err = validate_field_value(var_name, val_str, dropdown_options, dropdown_lookup_sets)
        if rule_err:
            new_rule_warns[internal_key] = rule_err
            if internal_key not in remaining_errors:
                remaining_errors.append(internal_key)

    key_note = f"key_note_{current_idx}_{note_key}"
    if key_note in st.session_state:
        st.session_state.df_data.at[current_idx, note_key] = st.session_state[key_note]

    if remaining_errors:
        st.session_state.err_reg[current_idx] = list(dict.fromkeys(remaining_errors))
        st.session_state.rule_warnings[current_idx] = new_rule_warns
        if current_idx in st.session_state.status_completed:
            st.session_state.status_completed.remove(current_idx)
        st.session_state.action_error_msg = (
            f"❌ Không thể chuyển người! Vẫn còn {len(remaining_errors)} trường bị lỗi hoặc chưa được sửa lại."
        )
        st.rerun()
        return

    st.session_state.err_reg.pop(current_idx, None)
    st.session_state.rule_warnings.pop(current_idx, None)
    st.session_state.status_completed.add(current_idx)
    st.session_state.action_error_msg = ""

    total_people = len(st.session_state.df_data)
    next_idx = None
    for i in range(current_idx + 1, total_people):
        if i not in st.session_state.status_completed:
            next_idx = i
            break
    if next_idx is None:
        for i in range(0, current_idx):
            if i not in st.session_state.status_completed:
                next_idx = i
                break

    st.session_state.person_selector_target = next_idx if next_idx is not None else current_idx
    st.rerun()


# ==========================================================================
# UI: SỬA LỖI
# ==========================================================================
if "df_data" in st.session_state:
    df_data = st.session_state.df_data
    err_reg = st.session_state.err_reg
    r_warns = st.session_state.get("rule_warnings", {})
    m_type = st.session_state.m_type
    dropdown_options = st.session_state.get("dropdown_options", {})
    field_meta = st.session_state.field_meta
    note_col_key = st.session_state.note_col_key

    st.success(
        f"✅ Tải file thành công! Mẫu: **"
        f"{'Trẻ dưới 6 tuổi' if m_type == 'UNDER6' else 'Trẻ 6-18 tuổi' if m_type == '6TO18' else 'Trên 18 tuổi'}"
        f"**. Tổng số dòng: **{len(df_data)}**"
    )

    if st.session_state.get("row_capacity_warning"):
        st.markdown(f"<div class='note-block'>{st.session_state.row_capacity_warning}</div>", unsafe_allow_html=True)

    if st.session_state.get("action_error_msg"):
        st.markdown(f"<div class='error-block'>{st.session_state.action_error_msg}</div>", unsafe_allow_html=True)

    if "person_selector_target" in st.session_state:
        st.session_state.person_selector_radio = st.session_state.person_selector_target
        del st.session_state.person_selector_target

    if "person_selector_radio" not in st.session_state:
        st.session_state.person_selector_radio = 0

    ho_ten_key = "HO_TEN" if "HO_TEN" in df_data.columns else None

    col_left, col_right = st.columns([1, 2.5])

    with col_left:
        st.subheader("📋 DANH SÁCH CẦN SỬA")

        def format_person(idx):
            name = df_data.iloc[idx].get(ho_ten_key, f"Dòng {idx+1}") if ho_ten_key else f"Dòng {idx+1}"
            if idx in st.session_state.status_completed:
                return f"🟢 {name} (Đã sửa xong)"
            n_err = len(err_reg.get(idx, []))
            if n_err == 0:
                return f"🟢 {name} (Đúng chuẩn)"
            return f"🔴 {name} ({n_err} lỗi)"

        selected_idx = st.radio(
            "Chọn người khám:",
            options=list(range(len(df_data))),
            format_func=format_person,
            key="person_selector_radio",
        )

    with col_right:
        curr_row = df_data.iloc[selected_idx]
        curr_person_name = curr_row.get(ho_ten_key, "") if ho_ten_key else ""

        is_completed = selected_idx in st.session_state.status_completed
        fields_to_edit = err_reg.get(selected_idx, [])
        person_rule_warns = r_warns.get(selected_idx, {})

        st.markdown(f"### 👤 Họ và tên: **{curr_person_name}**")

        if not fields_to_edit:
            st.info("🟢 Người khám này không có ô nào bị báo lỗi/tô màu. Dữ liệu hoàn toàn hợp lệ.")
        elif is_completed:
            st.info("🟢 Người khám này đã được đánh dấu **Đã sửa xong**.")
        else:
            st.warning(f"🔴 Người khám này đang còn **{len(fields_to_edit)}** trường cần xử lý lỗi.")

        if selected_idx in st.session_state.get("age_mismatch_rows", set()):
            st.markdown(
                "<div class='rule-warning'>⚠️ Tuổi của người khám không đúng với mẫu, "
                "nên dùng mẫu phù hợp với tuổi!</div>",
                unsafe_allow_html=True,
            )
        st.divider()

        active_fields = list(dict.fromkeys(fields_to_edit + list(person_rule_warns.keys())))

        for internal_key in active_fields:
            mf = field_meta[internal_key]
            var_name = mf["var_name"]
            label = mf["label"]
            key_input = f"key_input_{selected_idx}_{internal_key}"

            # [ICD-10] Áp dụng giá trị đã nối mã từ lượt bấm "➕ Thêm" trước đó
            # (nếu có) NGAY TẠI ĐÂY — tức là TRƯỚC khi widget key_input được
            # khởi tạo bên dưới. Bắt buộc phải làm theo thứ tự này: Streamlit
            # không cho phép ghi vào st.session_state[key_input] SAU khi widget
            # với key đó đã được khởi tạo trong cùng 1 lượt chạy (sẽ ném lỗi
            # StreamlitWidgetAlreadyInstantiatedError).
            pending_key = f"pending_{key_input}"
            if pending_key in st.session_state:
                st.session_state[key_input] = st.session_state.pop(pending_key)

            if key_input not in st.session_state:
                val_init = str(st.session_state.df_data.at[selected_idx, internal_key]).strip()
                st.session_state[key_input] = "" if val_init.lower() == "nan" else val_init

            if internal_key in person_rule_warns:
                st.markdown(
                    f"<div class='rule-warning'>⚠️ <b>Cảnh báo [{label}]:</b> {person_rule_warns[internal_key]}</div>",
                    unsafe_allow_html=True,
                )

            if var_name in MULTI_VALUE_ICD10_FIELDS:
                # [Đa giá trị ICD-10] Ô nhập liệu DUY NHẤT — nguồn dữ liệu thật,
                # gõ tay tự do, nhiều mã cách nhau bằng ';', đúng theo ví dụ mẫu
                # gốc "I10;A00;A01" (CHỈ MÃ THUẦN, không kèm tên bệnh).
                st.text_input(
                    f"**{label}** (nhiều mã cách nhau bằng ';', vd 'H52.1;J02.0;I10'):",
                    key=key_input,
                )

                # Tra cứu & nối mã — người dùng gõ để tìm (selectbox của Streamlit
                # hỗ trợ sẵn gõ-tìm-kiếm/autocomplete), chọn xong bấm "➕ Thêm" để
                # trích riêng phần MÃ (trước dấu " - ") và nối vào ô key_input phía
                # trên bằng ';', không trùng mã đã có, không dư ';' đầu/cuối. Việc
                # ghi giá trị + reset ô tra cứu CHỈ xảy ra khi bấm nút (st.button),
                # không tự kích hoạt theo lựa chọn của selectbox — tránh lặp lại sự
                # cố rerun tự phát đã từng xảy ra ở phiên bản multiselect trước đó.
                icd_options = dropdown_options.get(var_name, [])
                if icd_options:
                    gen_key = f"key_icd_gen_{selected_idx}_{internal_key}"
                    gen = st.session_state.get(gen_key, 0)
                    lookup_key = f"key_icd_lookup_{selected_idx}_{internal_key}_{gen}"

                    col_lookup, col_add = st.columns([5, 1])
                    with col_lookup:
                        picked_ref = st.selectbox(
                            f"🔍 Tra cứu & thêm mã ICD-10 cho {label} (gõ để tìm kiếm):",
                            options=[""] + icd_options,
                            key=lookup_key,
                        )
                    with col_add:
                        st.write("")
                        add_clicked = st.button(
                            "➕ Thêm", key=f"key_icd_addbtn_{selected_idx}_{internal_key}_{gen}"
                        )

                    if add_clicked:
                        if picked_ref:
                            ma_code = picked_ref.split(" - ", 1)[0].strip()
                            current_codes = [
                                c.strip()
                                for c in st.session_state[key_input].split(";")
                                if c.strip()
                            ]
                            if ma_code and ma_code not in current_codes:
                                current_codes.append(ma_code)
                                # KHÔNG ghi trực tiếp vào st.session_state[key_input]
                                # ở đây — widget key_input đã được khởi tạo phía trên
                                # trong CÙNG lượt chạy này (Streamlit sẽ báo lỗi
                                # StreamlitWidgetAlreadyInstantiatedError). Lưu tạm
                                # vào pending_key, áp dụng vào key_input ở ĐẦU lượt
                                # chạy tiếp theo (xem đoạn code trước khi khởi tạo
                                # key_input ở trên).
                                st.session_state[pending_key] = ";".join(current_codes)
                            # Đổi "thế hệ" key của ô tra cứu để nó tự làm mới/rỗng ở
                            # lần render tiếp theo, sẵn sàng cho lượt tìm mã kế tiếp.
                            st.session_state[gen_key] = gen + 1
                            st.rerun()
            elif var_name in dropdown_options:
                opts = dropdown_options[var_name]
                cur_val = st.session_state[key_input]
                full_opts = (
                    [cur_val] + [o for o in opts if o != cur_val]
                    if cur_val and cur_val not in opts else opts
                )
                sel_idx = full_opts.index(cur_val) if cur_val in full_opts else 0
                st.selectbox(f"**{label}** (Droplist chuẩn Master):", options=full_opts, index=sel_idx, key=key_input)
            else:
                st.text_input(f"**{label}:**", key=key_input)

        key_note = f"key_note_{selected_idx}_{note_col_key}"
        if key_note not in st.session_state:
            val_note_init = str(st.session_state.df_data.at[selected_idx, note_col_key]).strip()
            st.session_state[key_note] = "" if val_note_init.lower() == "nan" else val_note_init
        st.text_area(f"Cột {note_col_key} (không bắt buộc):", height=80, key=key_note)

        st.write("")
        col_btn1, col_btn2 = st.columns([2, 1])
        with col_btn1:
            if st.button("💾 Lưu & Chuyển người tiếp theo", type="primary", key=f"btn_save_main_{selected_idx}"):
                save_and_revalidate_person(selected_idx, active_fields, note_col_key)
        with col_btn2:
            if is_completed:
                if st.button("↩️ Đánh dấu chưa sửa", type="secondary", key=f"btn_unmark_{selected_idx}"):
                    st.session_state.status_completed.remove(selected_idx)
                    init_errs = st.session_state.get("initial_err_reg", {})
                    if selected_idx in init_errs:
                        st.session_state.err_reg[selected_idx] = list(init_errs[selected_idx])
                    st.session_state.person_selector_target = selected_idx
                    st.rerun()

    st.divider()

    col_act1, _ = st.columns([2, 1])
    with col_act1:
        if st.button("💾 TẠO FILE CHUẨN MASTERFLOW (_DASUA.xlsm)", type="primary"):
            if len(st.session_state.err_reg) > 0:
                st.error(
                    f"🚫 KHÔNG THỂ XUẤT FILE! Vẫn còn {len(st.session_state.err_reg)} người khám chưa được xử lý hết lỗi."
                )
            else:
                with st.spinner("Đang đóng gói dữ liệu và giữ nguyên cấu trúc VBA/Macro..."):
                    export_error = None
                    try:
                        master_path = st.session_state.master_path
                        wb_export = openpyxl.load_workbook(master_path, keep_vba=True)
                        sheet_name = [
                            s for s in wb_export.sheetnames
                            if s not in ["DM_CanLamSang", "Hướng dẫn"] and not s.startswith("dm")
                        ][0]
                        ws_export = wb_export[sheet_name]

                        # [FIX 2] Chỉ ghi vào đúng cột của các field thật (có col_idx
                        # xác định từ Master). KHÔNG đụng vào các cột công thức ẩn/
                        # STT (những cột này không nằm trong master_fields), nên
                        # công thức INDEX/MATCH tra tên tỉnh/xã được giữ nguyên.
                        for r_idx, df_row in st.session_state.df_data.iterrows():
                            excel_row = r_idx + 4
                            for mf in st.session_state.master_fields:
                                val = df_row.get(mf["internal_key"], "")
                                cell = ws_export.cell(row=excel_row, column=mf["col_idx"])
                                cell.value = val if val != "" else None
                                cell.fill = PatternFill(fill_type=None)

                        # [Vấn đề 1] Tô toàn bộ dòng của các trường hợp
                        # AGE_TEMPLATE_MISMATCH (NGAY_SINH hợp lệ nhưng lệch mẫu
                        # tuổi) bằng màu lavender, ghi đè sau cùng để không bị
                        # vòng lặp ở trên xoá mất.
                        lavender_fill = PatternFill(
                            start_color="D8D8F5", end_color="D8D8F5", fill_type="solid"
                        )
                        for r_idx in st.session_state.get("age_mismatch_rows", set()):
                            excel_row = r_idx + 4
                            for c in range(1, ws_export.max_column + 1):
                                ws_export.cell(row=excel_row, column=c).fill = lavender_fill

                        output = io.BytesIO()
                        wb_export.save(output)
                        output.seek(0)

                        clean_name = st.session_state.file_name.rsplit(".", 1)[0] + "_DASUA.xlsm"
                        st.session_state.export_file_ready = {"data": output, "name": clean_name}
                    except Exception as e:
                        # [Rà soát] Không để crash toàn app nếu Master bị khoá/hỏng/bị
                        # xoá lúc export — báo lỗi thân thiện, KHÔNG mất df_data đã sửa
                        # trong session, người dùng có thể bấm thử lại ngay.
                        export_error = str(e)

                if export_error:
                    st.error(
                        "🚫 KHÔNG THỂ TẠO FILE! Có lỗi khi ghi file Master (file có thể đang "
                        "mở ở chương trình khác, bị hỏng, hoặc bị di chuyển/xoá). Dữ liệu bạn "
                        f"đã sửa vẫn còn nguyên, hãy thử lại. Chi tiết: {export_error}"
                    )
                else:
                    st.rerun()

    if "export_file_ready" in st.session_state:
        st.success("🎉 Đã tạo file thành công! Bấm nút bên dưới để tải về hoặc chuyển sang file khác.")
        if st.session_state.get("row_capacity_warning"):
            st.markdown(f"<div class='note-block'>{st.session_state.row_capacity_warning}</div>", unsafe_allow_html=True)
        col_dl1, col_dl2 = st.columns([2, 1])
        with col_dl1:
            st.download_button(
                label="📥 TẢI FILE CHUẨN VỀ MÁY (_DASUA.xlsm)",
                data=st.session_state.export_file_ready["data"],
                file_name=st.session_state.export_file_ready["name"],
                mime="application/vnd.ms-excel.sheet.macroenabled.12",
            )
        with col_dl2:
            if st.button("➕ Tiếp tục sửa File khác", type="secondary"):
                reset_app_state()
