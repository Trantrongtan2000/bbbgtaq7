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

# System instruction: Hướng dẫn AI trả về JSON chuẩn
SYSTEM_INSTRUCTION = (
    "Bạn là một nhà phân tích tài liệu kỹ thuật. Nhiệm vụ của bạn là trích xuất thông tin từ 'Biên bản bàn giao' "
    "vào định dạng JSON. "
    "QUAN TRỌNG: Trường 'pk' (Phụ kiện) phải là một danh sách (Array) các chuỗi, không được gộp thành 1 chuỗi dài. "
    "Nếu không có thông tin, trả về null. Không thêm Markdown (```json)."
)

# Cấu hình file
CONFIG_FILE_PATH = 'config.ini'
TEMPLATE_FILE = 'bbbg.docx'

# --- CẤU HÌNH LỌC MODEL ---
DESIRED_MODELS_KEYWORDS = ['pro', 'flash']
EXCLUDE_MODELS_KEYWORDS = ['bison', 'gecko', 'embedding', 'aqa', 'vision', 'legacy']

# Tùy chỉnh file output
MAX_FILENAME_LEN = 200
MAX_SERI_DISPLAY = 100
MAX_DEVICES_IN_FILENAME = 2
DEFAULT_FONT_NAME = 'Times New Roman'
DEFAULT_FONT_SIZE = 12

# --- CÁC HÀM PHỤ TRỢ (HELPER FUNCTIONS) ---

def convert_none_to_empty_string(obj: Any) -> Any:
    """Đệ quy chuyển đổi các giá trị None thành chuỗi rỗng."""
    if isinstance(obj, dict):
        return {k: convert_none_to_empty_string(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [convert_none_to_empty_string(elem) for elem in obj]
    return "" if obj is None else obj

def clean_filename(filename: str) -> str:
    """Làm sạch tên file."""
    chars_to_remove = (r'[\\/*?":<>|.]')
    cleaned_name = re.sub(chars_to_remove, '', filename)
    if len(cleaned_name) > MAX_FILENAME_LEN:
        cleaned_name = cleaned_name[:MAX_FILENAME_LEN]
    return cleaned_name

def standardize_string(text: Any) -> str:
    """Chuẩn hóa chuỗi tiếng Việt."""
    if not isinstance(text, str):
        return str(text)
    
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
    """Rút gọn tên công ty."""
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

    for p in prefixes + suffixes:
        name_after_affix_removal = re.sub(r'^\s*' + re.escape(p) + r'\s*|' + r'\s*' + re.escape(p) + r'\s*$', '', name_after_affix_removal, flags=re.IGNORECASE).strip(" ,.-_&")

    name_after_common_removal = name_after_affix_removal
    for term in common_terms:
        name_after_common_removal = re.sub(r'\b' + re.escape(term) + r'\b', '', name_after_common_removal, flags=re.IGNORECASE).strip()
        name_after_common_removal = re.sub(r'\s+', ' ', name_after_common_removal).strip(" ,.-_&")

    if name_after_common_removal:
        return name_after_common_removal
    if name_after_affix_removal:
        return name_after_affix_removal
    return original_name

# --- CÁC HÀM XỬ LÝ LÕI (CORE LOGIC) ---

def group_devices(device_list: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Gộp các thiết bị giống hệt nhau."""
    grouped_devices = {}
    
    for item in device_list:
        if not isinstance(item, dict): continue
        
        # Xử lý 'pk' để làm key (vì list không hashable, phải chuyển về string)
        raw_pk = item.get('pk', '')
        if isinstance(raw_pk, list):
            pk_key = json.dumps(raw_pk, ensure_ascii=False, sort_keys=True)
        else:
            pk_key = str(raw_pk).strip()

        group_key_parts = [
            standardize_string(item.get('ttb', '')).strip(),
            str(item.get('model', '')).strip(),
            str(item.get('hang', '')).strip(),
            str(item.get('nsx', '')).strip(),
            str(item.get('dvt', '')).strip(),
            pk_key # Dùng chuỗi pk đã xử lý
        ]
        group_key = tuple(group_key_parts)

        # Xử lý số lượng (sl)
        current_sl_raw = item.get('sl', '0')
        try:
            cleaned_sl_str = re.sub(r'[^\d.]', '', str(current_sl_raw).strip())
            current_sl = float(cleaned_sl_str) if cleaned_sl_str else 0
        except (ValueError, TypeError):
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
                'pk_raw': raw_pk, # Lưu trữ giá trị gốc (list hoặc string)
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
            'pk': grouped_item['pk_raw'], # Trả về giá trị pk gốc
            'seri_text': seri_string
        })
    return final_device_list

def generate_filename(data: Dict[str, Any], grouped_devices: List[Dict[str, Any]]) -> str:
    """Tạo tên file Word."""
    device_parts = []
    for item in grouped_devices[:MAX_DEVICES_IN_FILENAME]:
        quantity = int(item.get('sl', 0))
        formatted_quantity = f"{quantity:02d}"
        device_name = str(item.get('ttb', '')).strip()
        cleaned_device_name = re.sub(r'[\\/*?":<>|{}\[\]().,_]', '', device_name).strip()
        if cleaned_device_name:
            device_parts.append(f"{formatted_quantity} {cleaned_device_name}")
    device_info_str = "-".join(device_parts) or "ThietBi"

    cty_name_full = str(data.get('cty', 'UnknownCompany')).strip()
    cleaned_cty_name = shorten_company_name(cty_name_full)
    if not cleaned_cty_name:
        cleaned_cty_name = re.sub(r'[\\/*?":<>|{}\[\]()]', '', cty_name_full).strip(" ,.-_&") or "CongTy"

    shd_value = str(data.get('shd', '')).strip()
    shd_main_part = shd_value.split('-', 1)[0].strip() or "SoDinhDanh"
    shd_cleaned = clean_filename(shd_main_part)

    raw_filename = f"{device_info_str}_{cleaned_cty_name}_{shd_cleaned}"
    final_filename_base = re.sub(r'\s+', '_', clean_filename(raw_filename)).strip('_')

    if not final_filename_base or len(final_filename_base) < 3:
        return f"BienBanBanGiao_{cleaned_cty_name}_{shd_cleaned}.docx"
        
    return final_filename_base + '.docx'

def fill_word_template(data: Dict[str, Any], grouped_devices: List[Dict[str, Any]]) -> BytesIO:
    """Điền dữ liệu vào Word (Xử lý phụ kiện thông minh)."""
    try:
        document = Document(TEMPLATE_FILE)
    except Exception as e:
        st.error(f"❌ Lỗi mở file mẫu '{TEMPLATE_FILE}'.", icon="❌")
        raise e

    # 1. ĐIỀN BẢNG
    try:
        table = document.tables[0]
        for i in range(len(table.rows) - 1, 0, -1):
            row = table.rows[i]
            row._element.getparent().remove(row._element)

        for count, item in enumerate(grouped_devices, 1):
            ttb_text = str(item.get('ttb', '')).strip()
            model_text = str(item.get('model', '')).strip()
            hang_text = str(item.get('hang', '')).strip()
            nsx_text = str(item.get('nsx', '')).strip()
            dvt_text = str(item.get('dvt', '')).strip()
            sl_text = str(int(item.get('sl', 0))).strip()
            
            # --- XỬ LÝ PHỤ KIỆN (CẢI TIẾN) ---
            raw_pk = item.get('pk', '')
            pk_lines = []

            # Nếu AI trả về List (nhờ prompt mới)
            if isinstance(raw_pk, list):
                pk_lines = [str(x).strip() for x in raw_pk if x]
            
            # Nếu AI trả về String (fallback)
            elif isinstance(raw_pk, str) and raw_pk:
                clean_str = re.sub(r'(cấu hình bao gồm|bao gồm|chi tiết cấu hình):', '', raw_pk, flags=re.IGNORECASE).strip()
                clean_str = clean_str.replace('–', '-').strip()
                # Tách bằng Dấu chấm phẩy (;) HOẶC Xuống dòng (\n)
                pk_lines = re.split(r'[;\n]+', clean_str)
            
            formatted_accessories = []
            for acc in pk_lines:
                clean_acc = acc.strip().lstrip('-•+').strip()
                if clean_acc:
                    formatted_accessories.append(f"  + {clean_acc}")
            
            device_info_text = f"{ttb_text}\n- Model: {model_text}\n- Hãng: {hang_text}\n- NSX: {nsx_text}"
            if formatted_accessories:
                device_info_text += "\n- Phụ kiện:\n" + "\n".join(formatted_accessories)
            # --------------------------------

            new_device_data = [str(count), device_info_text, dvt_text, sl_text, item['seri_text']]

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
        st.error("❌ File mẫu không có bảng.", icon="❌")
        raise

    # 2. REPLACE PLACEHOLDERS
    now = datetime.now()
    replacements = {
        "day": str(now.day),
        "month": str(now.month),
        "year": str(now.year),
    }

    shd_value = str(data.get('shd', '')).strip()
    shd_type = str(data.get('shd_type', 'Khác')).strip()
    if shd_value:
        shd_type_lower = standardize_string(shd_type)
        if any(x in shd_type_lower for x in ['hop dong', 'hd']):
            val = f"Dựa theo HĐ số: {shd_value}"
        elif any(x in shd_type_lower for x in ['po', 'de nghi']):
            val = f"Dựa theo PO: {shd_value}"
        else:
            val = f"Dựa theo số: {shd_value}"
        replacements["shd"] = val
    else:
        replacements["shd"] = ""

    shd_pattern = re.compile(re.escape("shd"), re.IGNORECASE)
    
    for p in document.paragraphs:
        # Replace date
        if any(x in p.text for x in ["day", "month", "year"]):
            for r in p.runs:
                for k, v in replacements.items():
                    if k in r.text:
                        r.text = r.text.replace(k, v)
        # Replace SHD
        if shd_pattern.search(p.text):
            for r in p.runs:
                if shd_pattern.search(r.text):
                    r.text = shd_pattern.sub(replacements["shd"], r.text)

    byte_io = BytesIO()
    document.save(byte_io)
    byte_io.seek(0)
    return byte_io

# --- API & CONFIG ---

@st.cache_resource
def check_prerequisites() -> bool:
    """Kiểm tra API key và file template."""
    if not os.path.exists(CONFIG_FILE_PATH):
        st.error(f"❌ Thiếu file '{CONFIG_FILE_PATH}'", icon="❌")
        return False
    
    try:
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE_PATH)
        api_key = config['API']['GEMINI_API_KEY']
        genai.configure(api_key=api_key)
    except Exception:
        st.error("❌ Lỗi đọc API Key.", icon="❌")
        return False

    if not os.path.exists(TEMPLATE_FILE):
        st.error(f"❌ Thiếu file mẫu '{TEMPLATE_FILE}'", icon="❌")
        return False
        
    return True

@st.cache_data
def get_filtered_models() -> List[str]:
    """Lấy và lọc model Gemini."""
    found = []
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                name = m.name.lower()
                if any(k in name for k in DESIRED_MODELS_KEYWORDS) and not any(k in name for k in EXCLUDE_MODELS_KEYWORDS):
                    found.append(m.name)
        
        # Sắp xếp ưu tiên: 2.5 > 2.0 > 1.5 Pro > Flash
        def priority(nm):
            n = nm.lower()
            if "gemini-3-pro-preview" in n: return 0
            if "gemini-2.5-pro" in n: return 1
            if "gemini-2.5-flash" in n: return 2
            if "gemini-2.5-flash-lite" in n: return 3
            return 4
            
        found.sort(key=priority)
        return found
    except Exception:
        return []

def call_gemini_vision_api(uploaded_file_part, prompt, model_list):
    """Gọi API với retry qua các model."""
    if not model_list:
        return None

    for model_name in model_list:
        try:
            with st.spinner(f"✨ Đang dùng model: {model_name}..."):
                model = genai.GenerativeModel(model_name=model_name, system_instruction=SYSTEM_INSTRUCTION)
                response = model.generate_content([uploaded_file_part, prompt])
                
                text = response.text.strip()
                # Clean Markdown json
                if text.startswith("```json"): text = text[7:]
                if text.endswith("```"): text = text[:-3]
                
                data = json.loads(text.strip())
                st.success(f"✅ Thành công với model: {model_name}")
                return data
        except Exception as e:
            print(f"Model {model_name} lỗi: {e}")
            continue
            
    return None

# --- MAIN ---

def main():
    st.set_page_config(page_title="Chuyển đổi Bàn giao", layout="centered")
    st.markdown("""<style>.stFileUploader {border: 1px dashed #004aad;}</style>""", unsafe_allow_html=True)
    st.title("Chuyển đổi Biên bản Bàn giao (Fix Lỗi Xuống dòng)")

    if not check_prerequisites():
        st.stop()

    available_models = get_filtered_models()
    if not available_models:
        st.error("Không tìm thấy model Gemini phù hợp.", icon="❌")
        st.stop()

    uploaded_file = st.file_uploader("Tải lên file (PDF/Ảnh)", type=["pdf", "jpg", "png"])

    if uploaded_file:
        st.info(f"📥 Đang xử lý: {uploaded_file.name}")
        
        file_bytes = uploaded_file.getvalue()
        mime = 'application/pdf' if uploaded_file.name.lower().endswith('.pdf') else 'image/jpeg'
        
        # --- PROMPT MỚI: YÊU CẦU PK LÀ MẢNG ---
        prompt_content = """
**Yêu cầu trích xuất JSON:**
1. **shd**: Số định danh.
2. **shd_type**: Loại (Hợp đồng, PO...).
3. **cty**: Tên công ty.
4. **ds**: Danh sách thiết bị:
   - **ttb**: Tên thiết bị
   - **model**: Model
   - **hang**: Hãng
   - **nsx**: Nước SX
   - **dvt**: ĐVT
   - **sl**: Số lượng
   - **seri**: Số seri
   - **pk**: QUAN TRỌNG - Trả về một DANH SÁCH (ARRAY) các chuỗi phụ kiện. 
     Ví dụ đúng: ["Dây nguồn (SL: 1)", "Cáp USB (SL: 1)"]
     Ví dụ sai: "Dây nguồn (SL: 1); Cáp USB (SL: 1)"

**Output JSON:**
{
  "shd": "", "shd_type": "", "cty": "",
  "ds": [
    { "ttb": "", "model": "", "hang": "", "nsx": "", "dvt": "", "sl": 0, "seri": null, "pk": [] }
  ]
}
"""
        
        data = call_gemini_vision_api({'mime_type': mime, 'data': file_bytes}, prompt_content, available_models)

        if data and 'ds' in data:
            data = convert_none_to_empty_string(data)
            grouped = group_devices(data['ds'])
            
            filename = generate_filename(data, grouped)
            word_io = fill_word_template(data, grouped)
            
            st.download_button("⬇️ Tải xuống file Word", word_io, filename, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.balloons()
        else:
            st.error("Không trích xuất được dữ liệu.", icon="❌")

if __name__ == "__main__":
    main()