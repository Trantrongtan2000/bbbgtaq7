import streamlit as st
import google.generativeai as genai
import os
import re
import json
import configparser
from io import BytesIO
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from typing import List, Dict, Any, Optional

# --- HẰNG SỐ (CONSTANTS) ---

# System instruction cho AI
SYSTEM_INSTRUCTION = (
    "Bạn là một nhà phân tích tài liệu kỹ thuật, chuyên trích xuất thông tin chi tiết từ 'Biên bản giao nhận - Nghiệm thu kiêm phiếu bảo hành' "
    "và các tài liệu tương tự. Nhiệm vụ của bạn là trích xuất các thông tin sau từ tệp PDF hoặc ảnh được cung cấp, đặc biệt là từ các bảng biểu, "
    "và **trả về DUY NHẤT dưới định dạng JSON hợp lệ**, không có bất kỳ văn bản giải thích, ký tự thừa, hoặc ký hiệu Markdown (như ```json) nào khác."
    "Sử dụng các viết tắt: shd (giá trị số định danh), shd_type (loại số định danh), cty, ds, ttb, model, hang, nsx, dvt, sl, seri, pk."
    "Lưu ý quan trọng: Nếu không tìm thấy Số seri hoặc Phụ kiện, hãy trả về giá trị là null cho các trường đó."
)

# Cấu hình
CONFIG_FILE_PATH = 'config.ini'
TEMPLATE_FILE = 'bbbg.docx'

# --- CẤU HÌNH LỌC MODEL ---
# Các từ khóa model bạn muốn ưu tiên (ví dụ: pro, flash)
DESIRED_MODELS_KEYWORDS = ['pro', 'flash']

# (CẬP NHẬT) Chỉ loại trừ các model cũ/chuyên dụng
EXCLUDE_MODELS_KEYWORDS = [
    'bison', 'gecko', 'embedding', 'aqa', 'vision', 'legacy'
    # Đã xóa '2.5-pro' khỏi danh sách này
]

# Tùy chỉnh file output
MAX_FILENAME_LEN = 200
MAX_SERI_DISPLAY = 100
MAX_DEVICES_IN_FILENAME = 2
DEFAULT_FONT_NAME = 'Times New Roman'
DEFAULT_FONT_SIZE = 12

# --- CÁC HÀM PHỤ TRỢ (HELPER FUNCTIONS) ---
# (Các hàm: convert_none_to_empty_string, clean_filename, 
# standardize_string, shorten_company_name giữ nguyên như cũ)

def convert_none_to_empty_string(obj: Any) -> Any:
    """Đệ quy chuyển đổi các giá trị None trong dicts và lists thành chuỗi rỗng."""
    if isinstance(obj, dict):
        return {k: convert_none_to_empty_string(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [convert_none_to_empty_string(elem) for elem in obj]
    return "" if obj is None else obj

def clean_filename(filename: str) -> str:
    """Loại bỏ các ký tự đặc biệt khỏi tên file và giới hạn độ dài."""
    chars_to_remove = (r'[\\/*?":<>|.]')
    cleaned_name = re.sub(chars_to_remove, '', filename)
    if len(cleaned_name) > MAX_FILENAME_LEN:
        cleaned_name = cleaned_name[:MAX_FILENAME_LEN]
    return cleaned_name

def standardize_string(text: Any) -> str:
    """Chuẩn hóa chuỗi tiếng Việt: loại bỏ dấu, lowercase, loại bỏ khoảng trắng thừa, dấu gạch ngang."""
    if not isinstance(text, str):
        return str(text)
    
    # Logic loại bỏ dấu (giữ nguyên)
    text = text.replace('Ằ', 'Ă').replace('Ắ', 'Ă').replace('Ặ', 'Ă').replace('Ẳ', 'Ă').replace('Ẵ', 'Ă')
    text = text.replace('È', 'E').replace('É', 'E').replace('Ẹ', 'E').replace('Ẻ', 'E').replace('Ẽ', 'E')
    text = text.replace('Ề', 'E').replace('Ế', 'E').replace('Ệ', 'E').replace('Ể', 'E').replace('Ễ', 'E')
    text = text.replace('Ì', 'I').replace('Í', 'I').replace('Ị', 'I').replace('Ỉ', 'I').replace('Ĩ', 'I')
    text = text.replace('Ò', 'O').replace('Ó', 'O').replace('Ọ', 'O').replace('Ỏ', 'O').replace('Õ', 'O')
    text = text.replace('Ồ', 'O').replace('Ố', 'O').replace('Ộ', 'O').replace('Ổ', 'O').replace('Ỗ', 'O')
    text = text.replace('Ờ', 'O').replace('Ớ', 'O').replace('Ợ', 'O').replace('Ở', 'O').replace('Ỡ', 'O')
    text = text.replace('Ù', 'U').replace('Ú', 'U').replace('Ụ', 'U').replace('Ủ', 'U').replace('Ũ', 'U')
    text = text.replace('Ừ', 'U').replace('Ứ', 'U').replace('Ự', 'U').replace('Ử', 'U').replace('Ữ', 'U')
    text = text.replace('Ỳ', 'Y').replace('Ý', 'Y').replace('Ỵ', 'Y').replace('Ỷ', 'Y').replace('Ỹ', 'Y')
    text = text.replace('Đ', 'D')
    
    text = text.lower()
    text = text.replace('-', ' ').strip()
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def shorten_company_name(company_name: str) -> str:
    """Rút gọn tên công ty (cải tiến logic fallback)."""
    if not isinstance(company_name, str):
        return str(company_name).strip()

    original_name = company_name.strip()
    name_after_affix_removal = original_name
    
    prefixes = [
        r"CÔNG TY TNHH MỘT THÀNH VIÊN", r"CÔNG TY TNHH MTV", r"CÔNG TY TNHH HAI THÀNH VIÊN TRỞ LÊN",
        r"CÔNG TY CỔ PHẦN", r"CÔNG TY TNHH", r"CÔNG TY", r"TNHH", r"CỔ PHẦN",
    ]
    suffixes = [
        r"MỘT THÀNH VIÊN", r"MTV", r"HAI THÀNH VIÊN TRỞ LÊN", r"CỔ PHẦN", r"TNHH",
    ]
    common_terms = [
        r"THƯƠNG MẠI VÀ DỊCH VỤ", r"DỊCH VỤ VÀ THƯƠNG MẠI", r"TM VÀ DV", r"DV VÀ TM", r"TM & DV", r"DV & TM",
        r"TM", r"DV", r"CÔNG NGHỆ", r"THƯƠNG MẠI", r"TRANG THIẾT BỊ", r"Y TẾ", r"XÂY DỰNG",
        r"ĐẦU TƯ", r"PHÁT TRIỂN", r"GIẢI PHÁP", r"KỸ THUẬT", r"SẢN XUẤT", r"NHẬP KHẨU", r"XUẤT NHẬP KHẨU",
        r"KINH DOANH", r"PHÂN PHỐI", r"VIỆT NAM"
    ]

    # 1. Loại bỏ tiền tố và hậu tố
    for p in prefixes + suffixes:
        name_after_affix_removal = re.sub(r'^\s*' + re.escape(p) + r'\s*|' + r'\s*' + re.escape(p) + r'\s*$', '', name_after_affix_removal, flags=re.IGNORECASE).strip(" ,.-_&")

    # 2. Loại bỏ các từ phổ biến
    name_after_common_removal = name_after_affix_removal
    for term in common_terms:
        name_after_common_removal = re.sub(r'\b' + re.escape(term) + r'\b', '', name_after_common_removal, flags=re.IGNORECASE).strip()
        name_after_common_removal = re.sub(r'\s+', ' ', name_after_common_removal).strip(" ,.-_&")

    # 3. Logic Fallback: Trả về kết quả tốt nhất có thể
    if name_after_common_removal:
        return name_after_common_removal
    if name_after_affix_removal:
        return name_after_affix_removal
    return original_name # Fallback an toàn nhất


# --- CÁC HÀM XỬ LÝ LÕI (CORE LOGIC FUNCTIONS) ---
# (Các hàm: group_devices, generate_filename, fill_word_template giữ nguyên)

def group_devices(device_list: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Gộp các thiết bị giống hệt nhau, tính tổng số lượng và gộp seri."""
    grouped_devices = {}
    
    for item in device_list:
        if not isinstance(item, dict): continue
        
        group_key_parts = [
            standardize_string(item.get('ttb', '')).strip(),
            str(item.get('model', '')).strip(),
            str(item.get('hang', '')).strip(),
            str(item.get('nsx', '')).strip(),
            str(item.get('dvt', '')).strip(),
            str(item.get('pk', '')).strip()
        ]
        group_key = tuple(group_key_parts)

        # Xử lý số lượng (sl)
        current_sl_raw = item.get('sl', '0')
        try:
            cleaned_sl_str = re.sub(r'[^\d.]', '', str(current_sl_raw).strip())
            current_sl = float(cleaned_sl_str) if cleaned_sl_str else 0
        except (ValueError, TypeError):
            print(f"Warning: Không thể chuyển đổi số lượng '{current_sl_raw}' thành số. Dùng giá trị 0.")
            current_sl = 0

        # Xử lý Seri
        current_seri = item.get('seri', [])
        if isinstance(current_seri, str):
            current_seri = [current_seri] if current_seri else []
        elif not isinstance(current_seri, list):
            current_seri = [str(current_seri)] if current_seri is not None else []

        cleaned_current_seri = [str(s).strip() for s in current_seri if s is not None and str(s).strip() != '']

        if group_key not in grouped_devices:
            grouped_devices[group_key] = {
                'ttb': str(item.get('ttb', '')).strip(),
                'model': str(item.get('model', '')).strip(),
                'hang': str(item.get('hang', '')).strip(),
                'nsx': str(item.get('nsx', '')).strip(),
                'dvt': str(item.get('dvt', '')).strip(),
                'pk': str(item.get('pk', '')).strip(),
                'total_sl': current_sl,
                'seri': set(cleaned_current_seri)
            }
        else:
            grouped_devices[group_key]['total_sl'] += current_sl
            grouped_devices[group_key]['seri'].update(cleaned_current_seri)

    # Chuyển đổi dictionary nhóm thành danh sách cuối cùng
    final_device_list = []
    for grouped_item in grouped_devices.values():
        seri_string = ""
        if grouped_item['seri']:
            unique_seri = sorted(list(grouped_item['seri']))
            display_seri = unique_seri[:MAX_SERI_DISPLAY]
            seri_string = 'Số seri: ' + ', '.join(display_seri)
            if len(unique_seri) > MAX_SERI_DISPLAY:
                seri_string += f" (và {len(unique_seri) - MAX_SERI_DISPLAY} seri khác)"
        
        final_device_list.append({
            'ttb': grouped_item['ttb'],
            'model': grouped_item['model'],
            'hang': grouped_item['hang'],
            'nsx': grouped_item['nsx'],
            'dvt': grouped_item['dvt'],
            'sl': grouped_item['total_sl'],
            'pk': grouped_item['pk'],
            'seri_text': seri_string
        })
    return final_device_list

def generate_filename(data: Dict[str, Any], grouped_devices: List[Dict[str, Any]]) -> str:
    """Tạo tên file Word đầu ra dựa trên dữ liệu."""
    
    # 1. Chuỗi thông tin thiết bị
    device_parts = []
    for item in grouped_devices[:MAX_DEVICES_IN_FILENAME]:
        quantity = int(item.get('sl', 0))
        formatted_quantity = f"{quantity:02d}"
        device_name = str(item.get('ttb', '')).strip()
        cleaned_device_name = re.sub(r'[\\/*?":<>|{}\[\]().,_]', '', device_name).strip()
        if cleaned_device_name:
            device_parts.append(f"{formatted_quantity} {cleaned_device_name}")
    device_info_str = "-".join(device_parts) or "ThietBi"

    # 2. Tên công ty
    cty_name_full = str(data.get('cty', 'UnknownCompany')).strip()
    cleaned_cty_name = shorten_company_name(cty_name_full)
    if not cleaned_cty_name:
        cleaned_cty_name = re.sub(r'[\\/*?":<>|{}\[\]()]', '', cty_name_full).strip(" ,.-_&") or "CongTy"

    # 3. SHD (Số định danh)
    shd_value = str(data.get('shd', '')).strip()
    shd_main_part = shd_value.split('-', 1)[0].strip() or "SoDinhDanh"
    shd_cleaned = clean_filename(shd_main_part)

    # 4. Kết hợp
    raw_filename = f"{device_info_str}_{cleaned_cty_name}_{shd_cleaned}"
    final_filename_base = re.sub(r'\s+', '_', clean_filename(raw_filename)).strip('_')

    if not final_filename_base or len(final_filename_base) < 3:
        return f"BienBanBanGiaoNoiBo_Fallback_{cleaned_cty_name}_{shd_cleaned}.docx"
        
    return final_filename_base + '.docx'

def fill_word_template(data: Dict[str, Any], grouped_devices: List[Dict[str, Any]]) -> BytesIO:
    """Điền dữ liệu vào file Word mẫu và trả về BytesIO."""
    
    try:
        document = Document(TEMPLATE_FILE)
    except Exception as e:
        st.error(f"❌ Không tìm thấy hoặc không mở được file mẫu '{TEMPLATE_FILE}'. Vui lòng đảm bảo file này nằm cùng thư mục.", icon="❌")
        raise e

    # --- 1. Điền bảng ---
    try:
        table = document.tables[0]
        # Xóa các hàng dữ liệu mẫu (trừ hàng tiêu đề đầu tiên)
        for i in range(len(table.rows) - 1, 0, -1):
            row = table.rows[i]
            row._element.getparent().remove(row._element)

        # Thêm hàng mới
        for count, item in enumerate(grouped_devices, 1):
            ttb_text = str(item.get('ttb', '')).strip()
            model_text = str(item.get('model', '')).strip()
            hang_text = str(item.get('hang', '')).strip()
            nsx_text = str(item.get('nsx', '')).strip()
            dvt_text = str(item.get('dvt', '')).strip()
            sl_text = str(int(item.get('sl', 0))).strip()
            pk_text = str(item.get('pk', '')).strip()

            device_info_text = f"{ttb_text}\n- Model: {model_text}\n- Hãng: {hang_text}\n- NSX: {nsx_text}"
            
            # Xử lý Phụ kiện (pk)
            if pk_text:
                pk_text = re.sub(r'(cấu hình bao gồm|bao gồm|chi tiết cấu hình):','', pk_text, flags=re.IGNORECASE).strip()
                pk_text = pk_text.replace('–', '-').strip()
                accessories = [f"  + {acc.strip().lstrip('-•').strip()}" for acc in pk_text.split('\n') if acc.strip()]
                if accessories:
                    device_info_text += "\n- Phụ kiện:\n" + "\n".join(accessories)

            new_device_data = [
                str(count),
                device_info_text,
                dvt_text,
                sl_text,
                item['seri_text']
            ]

            row = table.add_row()
            for i, cell_text in enumerate(new_device_data):
                ali = WD_ALIGN_PARAGRAPH.CENTER if i in (0, 2, 3) else WD_ALIGN_PARAGRAPH.LEFT
                cell = row.cells[i]
                cell.text = str(cell_text)
                for p in cell.paragraphs:
                    p.alignment = ali
                    for run in p.runs:
                        run.font.name = DEFAULT_FONT_NAME
                        run.font.size = Pt(DEFAULT_FONT_SIZE)

    except IndexError:
        st.error(f"❌ File mẫu '{TEMPLATE_FILE}' không chứa bảng nào.", icon="❌")
        raise
    except Exception as e:
        st.error(f"❌ Lỗi khi điền dữ liệu vào bảng: {e}", icon="❌")
        raise

    # --- 2. Thay thế placeholders (Ngày tháng, SHD) ---
    now = datetime.now()
    replacements = {
        "day": str(now.day),
        "month": str(now.month),
        "year": str(now.year),
    }

    # Định dạng SHD
    shd_value = str(data.get('shd', '')).strip()
    shd_type = str(data.get('shd_type', 'Khác')).strip()
    shd_value_to_replace = ""
    if shd_value:
        shd_type_lower = standardize_string(shd_type)
        if 'hop dong' in shd_type_lower or 'hd' in shd_type_lower:
            shd_value_to_replace = f"Dựa theo HĐ số: {shd_value}"
        elif 'po' in shd_type_lower or 'de nghi' in shd_type_lower:
            shd_value_to_replace = f"Dựa theo PO: {shd_value}"
        else:
            shd_value_to_replace = f"Dựa theo số: {shd_value}"
    
    replacements["shd"] = shd_value_to_replace
    print(f"Giá trị thay thế cho 'shd': '{shd_value_to_replace}'")

    # Thực hiện thay thế
    shd_placeholder_found = False
    shd_pattern = re.compile(re.escape("shd"), re.IGNORECASE)
    
    for p in document.paragraphs:
        # Thay thế ngày tháng
        if "day" in p.text or "month" in p.text or "year" in p.text:
            for r in p.runs:
                for key, val in replacements.items():
                    if key in r.text:
                        r.text = r.text.replace(key, val)
        
        # Thay thế SHD
        if shd_pattern.search(p.text):
            for r in p.runs:
                if shd_pattern.search(r.text):
                    r.text = shd_pattern.sub(shd_value_to_replace, r.text)
                    shd_placeholder_found = True

    if not shd_placeholder_found:
        st.warning("⚠️ Không tìm thấy placeholder 'shd' trong file mẫu. Số HĐ/PO sẽ không được điền.", icon="⚠️")

    # --- 3. Lưu vào BytesIO ---
    byte_io = BytesIO()
    document.save(byte_io)
    byte_io.seek(0)
    return byte_io


# --- CÁC HÀM TƯƠNG TÁC API (API & CONFIG) ---

@st.cache_resource
def check_prerequisites() -> bool:
    """Kiểm tra API key và file template. Trả về True nếu tất cả đều OK."""
    # 1. Kiểm tra API Key
    api_key_ok = False
    if not os.path.exists(CONFIG_FILE_PATH):
        st.error(f"❌ Lỗi cấu hình: Không tìm thấy file '{CONFIG_FILE_PATH}'.", icon="❌")
    else:
        try:
            config = configparser.ConfigParser()
            config.read(CONFIG_FILE_PATH)
            api_key = config['API']['GEMINI_API_KEY']
            genai.configure(api_key=api_key)
            print("Đã đọc API Key và cấu hình genai.")
            api_key_ok = True
        except Exception as e:
            st.error(f"❌ Lỗi cấu hình: Không đọc được API Key từ '{CONFIG_FILE_PATH}': {e}", icon="❌")

    # 2. Kiểm tra file mẫu Word
    template_ok = os.path.exists(TEMPLATE_FILE)
    if not template_ok:
        st.error(f"❌ Lỗi file mẫu: Không tìm thấy file '{TEMPLATE_FILE}'. Vui lòng đảm bảo file này nằm cùng thư mục.", icon="❌")

    return api_key_ok and template_ok

@st.cache_data
def get_filtered_models() -> List[str]:
    """
    (CẬP NHẬT) Lấy danh sách model từ API và lọc ra các model phù hợp.
    Ưu tiên các model Pro mạnh nhất (2.5 -> 2.0 -> 1.5).
    """
    print("Đang truy vấn danh sách model từ API...")
    found_models = []
    try:
        for m in genai.list_models():
            # Kiểm tra xem model có hỗ trợ 'generateContent' không
            if 'generateContent' in m.supported_generation_methods:
                model_name = m.name.lower() # Ví dụ: 'models/gemini-1.5-pro-latest'
                
                # 1. Lọc
                has_desired = any(k in model_name for k in DESIRED_MODELS_KEYWORDS)
                has_excluded = any(k in model_name for k in EXCLUDE_MODELS_KEYWORDS)
                
                # Chỉ thêm vào danh sách nếu nó chứa từ khóa mong muốn (pro, flash)
                # VÀ KHÔNG chứa từ khóa loại trừ (bison, gecko, ...)
                if has_desired and not has_excluded:
                    found_models.append(m.name) # Thêm tên model đầy đủ
        
        # (MỚI) Logic sắp xếp ưu tiên (Pro 2.5 -> 2.0 -> 1.5 -> Flash 1.5)
        def get_priority(model_name):
            """Hàm trả về tuple để sort, số nhỏ hơn = ưu tiên cao hơn."""
            name = model_name.lower()
            
            # Ưu tiên 1: Gemini 2.5 Pro (preview, latest, v.v.)
            if "gemini-2.5-pro" in name:
                return (0, "preview" not in name, name) # Ưu tiên preview (0, False, ...)
            
            # Ưu tiên 2: Gemini 2.0 Pro (nếu có)
            if "gemini-2.0-pro" in name:
                return (1, "latest" not in name, name) # Ưu tiên latest
            
            # Ưu tiên 3: Gemini 1.5 Pro
            if "gemini-1.5-pro-latest" in name:
                return (2, name)
            if "gemini-1.5-pro" in name:
                return (3, name)
            
            # Ưu tiên 4: Gemini 1.5 Flash
            if "gemini-1.5-flash-latest" in name:
                return (4, name)
            if "gemini-1.5-flash" in name:
                return (5, name)
            
            # Mặc định
            return (6, name)

        # Sắp xếp danh sách model dựa trên hàm ưu tiên
        found_models.sort(key=get_priority)
        
        print(f"Đã lọc và sắp xếp (Ưu tiên Pro 2.5) {len(found_models)} model. Thứ tự ưu tiên: {found_models}")
        
        if not found_models:
             print("Không tìm thấy model nào. Đảm bảo DESIRED/EXCLUDE_KEYWORDS đúng.")
             st.warning("Không tìm thấy model nào phù hợp sau khi lọc. Vui lòng kiểm tra lại cấu hình lọc.")
             
        return found_models

    except Exception as e:
        st.error(f"Lỗi khi truy vấn danh sách model: {e}", icon="❌")
        print(f"Lỗi khi gọi genai.list_models(): {e}")
        return [] # Trả về danh sách rỗng nếu lỗi

def call_gemini_vision_api(
    uploaded_file_part: Dict[str, Any], 
    prompt: str,
    model_list: List[str]
) -> Optional[Dict[str, Any]]:
    """
    (CẬP NHẬT) Gọi API Gemini với danh sách model đã lọc, trả về dict JSON hoặc None.
    """
    data = None
    raw_ai_response = ""
    
    if not model_list:
        st.error("❌ Không tìm thấy model nào phù hợp để xử lý. Vui lòng kiểm tra lại cấu hình lọc model.", icon="❌")
        return None

    for model_name in model_list:
        try:
            with st.spinner(f"✨ Đang trích xuất dữ liệu bằng model: **{model_name}**..."):
                model = genai.GenerativeModel(
                    model_name=model_name,
                    system_instruction=SYSTEM_INSTRUCTION
                )
                response = model.generate_content(
                    contents=[uploaded_file_part, prompt]
                )
                raw_ai_response = response.text
                print(f"Phản hồi thô từ {model_name}: {raw_ai_response}")

                # Làm sạch và parse JSON
                cleaned_response = raw_ai_response.strip().removeprefix("```json").removesuffix("```").strip()
                data = json.loads(cleaned_response)
                
                st.success(f"Trích xuất thành công bằng model: **{model_name}**!")
                return data # Thành công, trả về dữ liệu
                
        except json.JSONDecodeError as json_err:
            st.warning(f"⚠️ Model {model_name} không trả về JSON hợp lệ: {json_err}. Đang thử model tiếp theo...", icon="⚠️")
            print(f"Model {model_name} lỗi JSON: {json_err}")
        except Exception as api_err:
            # Kiểm tra xem có phải lỗi Quota 429 không
            if "429" in str(api_err) and "quota" in str(api_err).lower():
                 st.warning(f"⚠️ Model {model_name} báo lỗi Quota (429). Đang thử model tiếp theo...", icon="⚠️")
                 print(f"Model {model_name} lỗi Quota 429.")
            else:
                # Lỗi API khác
                st.warning(f"⚠️ Model {model_name} gặp lỗi API: {api_err}. Đang thử model tiếp theo...", icon="⚠️")
                print(f"Model {model_name} lỗi API: {api_err}")

    # Nếu vòng lặp kết thúc mà không thành công
    st.error("❌ Tất cả các model đã thử đều thất bại.", icon="❌")
    if raw_ai_response:
        st.text_area("Phản hồi gốc cuối cùng (gây lỗi):", raw_ai_response, height=200)
    return None

# --- HÀM CHÍNH (MAIN FUNCTION) ---

def main():
    st.set_page_config(page_title="Chuyển đổi Bàn giao", layout="centered")
    
    # --- CSS Tùy chỉnh ---
    st.markdown("""
    <style>
    .stFileUploader {
        padding: 1rem;
        border: 1px dashed #004aad;
        border-radius: 0.5rem;
        background-color: rgba(230, 240, 255, 0.1);
        margin-bottom: 1.5rem;
    }
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    div[data-testid="stVerticalBlock"] {
        gap: 1.5rem;
    }
    h1 {
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)

    st.title("Công cụ Chuyển đổi Biên bản Bàn giao")

    # --- 1. KIỂM TRA ĐIỀU KIỆN TIÊN QUYẾT ---
    if not check_prerequisites():
        st.warning("Vui lòng khắc phục các lỗi cấu hình trên để tiếp tục.", icon="⚠️")
        st.stop() # Dừng ứng dụng nếu thiếu API key hoặc file mẫu

    st.markdown(f"ℹ️ **Lưu ý:** File mẫu Word (`{TEMPLATE_FILE}`) đã được tìm thấy.")
    
    # --- (MỚI) LẤY DANH SÁCH MODEL SAU KHI QUA BƯỚC KIỂM TRA ---
    available_models = get_filtered_models()
    if not available_models:
        st.error("Không thể lấy được danh sách model phù hợp từ Google. Vui lòng kiểm tra API key và kết nối.", icon="❌")
        st.stop()

    # --- 2. GIAO DIỆN TẢI LÊN ---
    st.subheader("Tải lên Biên bản bàn giao gốc (PDF hoặc Ảnh)")
    uploaded_file = st.file_uploader(
        "Chọn file Biên bản bàn giao công ty (PDF hoặc Ảnh)",
        type=["pdf", "jpg", "jpeg", "png"],
        label_visibility="collapsed",
        key="file_uploader"
    )

    if not uploaded_file:
        st.info("⬆️ Vui lòng chọn một file PDF/Ảnh để bắt đầu.", icon="📄")
        st.stop()

    # --- 3. XỬ LÝ FILE (CHỈ CHẠY KHI CÓ FILE) ---
    try:
        st.info(f"📥 Đang xử lý file: **{uploaded_file.name}**", icon="⏳")
        
        # 3.1. Đọc file và chuẩn bị
        file_bytes = uploaded_file.getvalue()
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension == 'pdf':
            file_mime_type = 'application/pdf'
        elif file_extension in ['jpg', 'jpeg']:
            file_mime_type = 'image/jpeg'
        elif file_extension == 'png':
            file_mime_type = 'image/png'
        else:
            st.error("Định dạng file không được hỗ trợ.", icon="❌")
            st.stop()

        uploaded_file_part = {
            'mime_type': file_mime_type,
            'data': file_bytes
        }

        # 3.2. Prompt cho AI
        prompt_content = """
**Thông tin cần trích xuất:**
- **Số định danh chính (shd):** Giá trị số hoặc mã của biên bản.
- **Loại số định danh (shd_type):** Xác định loại của 'shd'.
- **Tên công ty bàn giao (cty):** Tên đầy đủ của công ty bên giao (Bên A).
- **Danh sách thiết bị (ds):** Mảng các đối tượng JSON (ttb, model, hang, nsx, dvt, sl, seri, pk).

**Cấu trúc JSON đầu ra phải tuân thủ nghiêm ngặt:**
{
  "shd": "Giá trị số/mã",
  "shd_type": "Hợp đồng" hoặc "PO" hoặc "Đề nghị" hoặc "Khác",
  "cty": "Tên công ty",
  "ds": [
    {
      "ttb": "Tên thiết bị",
      "model": "Model thiết bị",
      "hang": "Hãng sản xuất",
      "nsx": "Nước sản xuất",
      "dvt": "Đơn vị tính",
      "sl": "Số lượng",
      "seri": "Số seri" hoặc ["seri1", "seri2"] hoặc null,
      "pk": "Chi tiết phụ kiện" hoặc null
    }
  ]
}
"""

        # 3.3. Gọi API Gemini (với danh sách model đã lọc)
        data = call_gemini_vision_api(uploaded_file_part, prompt_content, available_models)

        if not data:
            st.error("Không thể trích xuất dữ liệu từ file.", icon="❌")
            st.stop()

        # 3.4. Kiểm tra dữ liệu AI trả về
        if 'ds' not in data or not isinstance(data.get('ds'), list):
            st.error("❌ Phản hồi từ AI không chứa danh sách thiết bị ('ds') hợp lệ.", icon="❌")
            st.text_area("Dữ liệu AI trả về:", json.dumps(data, indent=2, ensure_ascii=False), height=200)
            st.stop()

        # 3.5. Xử lý và tạo file Word
        st.info("✍️ Đang tạo file Word...", icon="⏳")
        
        # Chuyển đổi None -> ""
        data = convert_none_to_empty_string(data)
        
        # Gộp nhóm thiết bị
        grouped_devices = group_devices(data['ds'])
        if not grouped_devices:
            st.warning("⚠️ Không có thiết bị hợp lệ nào được tìm thấy sau khi gộp nhóm.", icon="⚠️")
            st.stop()

        # Tạo tên file
        filename = generate_filename(data, grouped_devices)
        
        # Điền vào file Word
        word_bytes_io = fill_word_template(data, grouped_devices)

        # 3.6. Hiển thị nút Tải xuống
        st.download_button(
            label=f"✅ Tải xuống file: {filename}",
            data=word_bytes_io,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.success(f"🎉 Đã tạo file thành công: **{filename}**", icon="✅")

    except Exception as e:
        st.error(f"❌ Đã có lỗi không mong muốn xảy ra trong quá trình xử lý: {e}", icon="❌")
        print(f"Lỗi không mong muốn trong hàm main: {e}")

if __name__ == "__main__":
    main()