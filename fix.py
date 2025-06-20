import google.generativeai as genai
import streamlit as st
import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import tempfile
import json
import re
import configparser
from io import BytesIO

def convert_none_to_empty_string(obj):
    """Recursively converts None values in dictionaries and lists to empty strings."""
    if isinstance(obj, dict):
        return {k: convert_none_to_empty_string(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_none_to_empty_string(elem) for elem in obj]
    elif obj is None:
        return ""
    else:
        return obj

def clean_filename(filename):
    """Loại bỏ các ký tự đặc biệt khỏi tên file và giới hạn độ dài."""
    chars_to_remove = (r'[\\/*?":<>|.]')
    cleaned_name = re.sub(chars_to_remove, '', filename)
    max_len = 200 # Giới hạn độ dài tên file
    if len(cleaned_name) > max_len:
        cleaned_name = cleaned_name[:max_len]
    return cleaned_name

# --- Hàm chuẩn hóa ký tự tiếng Việt và làm sạch cho grouping/filename ---
def standardize_string(text):
    """Chuẩn hóa chuỗi: loại bỏ dấu, chuyển lowercase, loại bỏ khoảng trắng thừa, dấu gạch ngang."""
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
    text = re.sub(r'\s+', ' ', text).strip()
    text = text.replace('-', ' ').strip()
    text = re.sub(r'\s+', ' ', text).strip()

    return text

# --- Hàm rút gọn tên công ty ---
def shorten_company_name(company_name):
    """Rút gọn tên công ty bằng cách loại bỏ các tiền tố và hậu tố phổ biến."""
    if not isinstance(company_name, str):
        return str(company_name).strip()

    cleaned_name = company_name.strip()
    upper_name = cleaned_name.upper()

    prefixes = [
        "CÔNG TY TNHH MỘT THÀNH VIÊN", "CÔNG TY TNHH MTV", "CÔNG TY TNHH HAI THÀNH VIÊN TRỞ LÊN",
        "CÔNG TY CỔ PHẦN", "CÔNG TY TNHH", "CÔNG TY", "TNHH", "CỔ PHẦN",
    ]
    suffixes = [
        "MỘT THÀNH VIÊN", "MTV", "HAI THÀNH VIÊN TRỞ LÊN", "CỔ PHẦN", "TNHH",
    ]
    common_terms = [
        "THƯƠNG MẠI VÀ DỊCH VỤ", "DỊCH VỤ VÀ THƯƠNG MẠI", "TM VÀ DV", "DV VÀ TM", "TM & DV", "DV & TM",
        "TM", "DV", "CÔNG NGHỆ", "THƯƠNG MẠI", "TRANG THIẾT BỊ", "Y TẾ", "XÂY DỰNG",
        "ĐẦU TƯ", "PHÁT TRIỂN", "GIẢI PHÁP", "KỸ THUẬT", "SẢN XUẤT", "NHẬP KHẨU", "XUẤT NHẬP KHẨU",
        "KINH DOANH", "PHÂN PHỐI", "VIỆT NAM"
    ]

    for prefix in prefixes:
        pattern = r'^' + re.escape(prefix) + r'\s*'
        cleaned_name = re.sub(pattern, '', cleaned_name, flags=re.IGNORECASE).strip(" ,.-_&")

    for suffix in suffixes:
         pattern = r'\s*' + re.escape(suffix) + r'$'
         cleaned_name = re.sub(pattern, '', cleaned_name, flags=re.IGNORECASE).strip(" ,.-_&")

    for term in common_terms:
         pattern = r'\b' + re.escape(term) + r'\b'
         cleaned_name = re.sub(pattern, '', cleaned_name, flags=re.IGNORECASE).strip()
         cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()

    cleaned_name = cleaned_name.strip(" ,.-_&")

    if not cleaned_name:
        words = company_name.strip().split()
        if words:
             fallback_name = " ".join(words[-3:])
             return fallback_name.strip(" ,.-_&")

        return company_name.strip()

    return cleaned_name

# --- Kết thúc hàm rút gọn tên công ty ---

# --- Cấu hình giao diện và CSS ---
# Đổi layout từ wide sang centered
st.set_page_config(page_title="Chuyển đổi Bàn giao", layout="centered")

st.markdown("""
<style>
/* Loại bỏ màu nền tùy chỉnh để sử dụng màu nền mặc định của Streamlit Theme (Dark/Light) */
/* .stApp { background-color: #f0f2f6; } */

/* Loại bỏ padding ngang tùy chỉnh để Streamlit centered layout quản lý */
.css-1lcbmhc { /* Đây là class cho main content container */
    padding-top: 0rem;
    padding-bottom: 10rem;
    /* Loại bỏ padding-left và padding-right */
    /* padding-left: 5%; */
    /* padding-right: 5%; */
}

/* Loại bỏ màu chữ tùy chỉnh để sử dụng màu chữ mặc định của Streamlit Theme */
/* .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { color: #0e1117; } */

/* Kiểu dáng cho File Uploader */
.stFileUploader {
    padding: 1rem;
    border: 1px dashed #004aad; /* Viền nét đứt */
    border-radius: 0.5rem;
    background-color: rgba(230, 240, 255, 0.1); /* Nền nhẹ có độ trong suốt, phù hợp với nền tối */
    margin-bottom: 1.5rem;
}

/* Màu cho Progress Bar */
.stProgress > div > div > div > div {
    background-color: #4CAF50; /* Màu xanh lá */
}

/* Khoảng cách giữa các Block */
div[data-testid="stVerticalBlock"] {
    gap: 1.5rem;
}

/* Padding cho block container (có thể thừa nếu .css-1lcbmhc đã xử lý) */
.reportview-container .main .block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}

/* Căn giữa tiêu đề */
h1 {
    text-align: center;
}


</style>
""", unsafe_allow_html=True)
# --- Kết thúc Cấu hình giao diện và CSS ---

# --- Cấu hình Google API Key từ config.ini ---
# Đặt khối này ở đầu script để đảm bảo API được cấu hình sớm nhất có thể
config = configparser.ConfigParser()
config_file_path = 'config.ini'
google_api_key = None # Khởi tạo biến

if os.path.exists(config_file_path):
    config.read(config_file_path)
    try:
        # Dòng này đọc giá trị của GEMINI_API_KEY từ section [API] trong file config.ini
        google_api_key = config['API']['GEMINI_API_KEY'] # <-- DÒNG LẤY API KEY

        # Dòng này sử dụng giá trị API key vừa đọc để cấu hình thư viện Google Generative AI
        genai.configure(api_key=google_api_key)      # <-- DÒNG SỬ DỤNG API KEY ĐỂ CẤU HÌNH

        print("Đã đọc API Key từ config.ini và cấu hình genai.")

    except KeyError:
        st.error(f"Lỗi cấu hình: File '{config_file_path}' không có section [API] hoặc key GEMINI_API_KEY. Vui lòng kiểm tra lại file config.ini.", icon="❌")
        # Không st.stop() ngay đây để hiển thị giao diện còn lại, nhưng sẽ báo lỗi khi gọi API
        google_api_key = None # Đảm bảo key là None nếu có lỗi đọc config
    except Exception as e:
        st.error(f"Lỗi khi đọc file cấu hình '{config_file_path}': {e}. Vui lòng kiểm tra định dạng file config.ini.", icon="❌")
        # Không st.stop() ngay đây
        google_api_key = None # Đảm bảo key là None nếu có lỗi khác
else:
    st.error(f"Lỗi cấu hình: Không tìm thấy file cấu hình '{config_file_path}'. Vui lòng tạo file này với section [API] và key GEMINI_API_KEY.", icon="❌")
    # Không st.stop() ngay đây
    google_api_key = None # Đảm bảo key là None nếu không tìm thấy file
# --- Kết thúc cấu hình API Key ---


st.title("Công cụ Chuyển đổi Biên bản Bàn giao")
# Sử dụng cột để bố cục phần upload và thông tin (vẫn giữ cột để tổ chức)
# col1, col2 = st.columns([2, 1]) # Tỷ lệ cột - Có thể bỏ cột nếu muốn đơn giản hơn trong centered layout

# with col1: # Nếu bỏ cột, đưa nội dung ra ngoài with block
st.subheader("Tải lên Biên bản bàn giao gốc (PDF hoặc Ảnh)")
file_name = st.file_uploader("Chọn file Biên bản bàn giao công ty (PDF hoặc Ảnh)", type=["pdf", "jpg", "jpeg", "png"], label_visibility="collapsed", key="file_uploader")

# with col2: # Nếu bỏ cột, đưa nội dung ra ngoài with block
st.markdown("ℹ️ **Lưu ý:** File mẫu Word (`bbbg.docx`) phải nằm cùng thư mục với script.")


temp_file_path = None

# Chỉ tiếp tục xử lý nếu có file được tải lên VÀ API Key đã được cấu hình thành công
if file_name is not None and google_api_key is not None:
    try:
        st.info(f"📥 Đang tải lên và xử lý file: **{file_name.name}**", icon="⏳")
        # Determine suffix based on file type
        file_extension = file_name.name.split('.')[-1].lower()
        suffix = f".{file_extension}" if file_extension in ["pdf", "jpg", "jpeg", "png"] else ""

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
            temp_file.write(file_name.getvalue())
            temp_file_path = temp_file.name
            print(f"File tạm được lưu tại: {temp_file_path}")

        # --- Chọn Model AI ---
        model_name_list = ["gemini-1.5-flash", "gemini-1.5-pro"] # Danh sách các model có thể dùng
        current_model = model_name_list[0] # Chọn model mặc định (ví dụ: flash)

        with st.spinner("✨ Đang trích xuất dữ liệu từ file PDF..."):

            # Tạo đối tượng model sau khi API Key đã được cấu hình
            model = genai.GenerativeModel(
                # Sử dụng model_name đã được chọn
                model_name='gemini-2.0-flash-lite',
                system_instruction=[
                    "Bạn là một nhà phân tích tài liệu kỹ thuật, chuyên trích xuất thông tin chi tiết từ 'Biên bản giao nhận - Nghiệm thu kiêm phiếu bảo hành' và các tài liệu tương tự.",
                    "Nhiệm vụ của bạn là trích xuất các thông tin sau từ tệp PDF hoặc ảnh được cung cấp, đặc biệt là từ các bảng biểu, và trả về dưới định dạng JSON hợp lệ.",
                    "Đảm bảo đầu ra JSON tuân thủ cấu trúc được yêu cầu trong prompt, sử dụng các viết tắt: shd (cho giá trị số định danh), shd_type (cho loại số định danh), cty, ds, ttb, model, hang, nsx, dvt, sl, seri, pk."
                ],
            )

            sample_pdf = genai.upload_file(path=temp_file_path)
            print(f"File đã tải lên Google AI: {sample_pdf.name}")

            prompt ="""
**Thông tin cần trích xuất:**
- **Số định danh chính (shd):** Giá trị số hoặc mã của biên bản (ví dụ: Số hợp đồng, số đề xuất, mã đề nghị, số PO). Chỉ xuất hiện một lần.
- **Loại số định danh (shd_type):** Xác định loại của 'shd' (ví dụ: "Hợp đồng", "PO", "Đề nghị", "Khác"). Dựa vào các cụm từ như "HĐ số:", "Theo HĐ số:", "Số Hợp Đồng:", "Dựa theo HĐ số:", "PO số:", "Số PO:", "Dựa theo số PO:", "Mã đề nghị:", "Số đề xuất:". Nếu không rõ loại, dùng "Khác". Chỉ xuất hiện một lần.
- **Tên công ty bàn giao (cty):** Tên đầy đủ của công ty bên giao (Bên A). Chỉ xuất hiện một lần.
- **Danh sách thiết bị (ds):** Một mảng các đối tượng JSON, mỗi đối tượng đại diện cho một thiết bị được trích xuất từ bảng. Đối với mỗi thiết bị, hãy xác định:
    - **Tên thiết bị (ttb):** Dựa vào cột 'MÔ TẢ'.
    - **Model (model):**
    - **Hãng (hang):**
    - **Nước sản xuất (nsx):**
    - **Đơn vị tính (dvt):**
    - **Số lượng (sl):**
    - **Số seri (seri):** Dựa vào cột 'IMEI'. Nếu có nhiều seri cho một dòng thiết bị, trả về dưới dạng mảng chuỗi. Nếu chỉ có một, trả về chuỗi. Nếu không có, trả về `null`.
    - **Phụ kiện (pk):** Chi tiết phụ kiện hoặc cấu hình kỹ thuật. Dữ liệu dạng chuỗi. Nếu có nhiều dòng phụ kiện cho một thiết bị, nối lại và xuống dòng bằng `\n`. Nếu không có, trả về `null`.

**Cấu trúc JSON mong muốn:**
```json
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
      "seri": "Số seri" hoặc ["seri1", "seri2"] hoặc "",
      "pk": "Gồm:\n- Phụ kiện A (SL: ... ĐVT: ...)\n- Phụ kiện B..." hoặc ""
    }
  ]
}
```
Đảm bảo đầu ra là JSON hợp lệ và chỉ chứa dữ liệu JSON, không có bất kỳ văn bản giải thích nào khác.
"""
            response = model.generate_content([sample_pdf, prompt])

            try:
                 sample_pdf.delete()
                 print(f"File đã xóa trên Google AI: {sample_pdf.name}")
            except Exception as e:
                 print(f"Lỗi khi xóa file trên Google AI: {e}")

        # --- Xử lý phản hồi từ AI ---
        a = response.text.strip()
        if a.startswith("```json"):
            a = a[len("```json"):].strip()
        if a.endswith("```"):
             a = a[:-len("```")].strip()

        data = None
        try:
            data = json.loads(a)
            # Convert all None values to empty strings recursively
            data = convert_none_to_empty_string(data)
            print("Dữ liệu JSON nhận được (sau khi xử lý None):", json.dumps(data, indent=2, ensure_ascii=False))

            extracted_shd = data.get('shd')
            extracted_shd_type = data.get('shd_type')
            print(f"Extracted shd value from AI: '{extracted_shd}' (Type: '{extracted_shd_type}')")

            if 'ds' not in data or not isinstance(data.get('ds'), list):
                 st.error("❌ Phản hồi từ AI không chứa danh sách thiết bị hợp lệ ('ds'). Vui lòng thử lại với file khác hoặc kiểm tra nội dung file.", icon="❌")
                 print(f"Phản hồi AI thiếu khóa 'ds' hoặc 'ds' không phải list: {data}")
                 data = None
            if data and 'shd' not in data:
                 print(f"Phản hồi AI thiếu khóa 'shd', gán giá trị mặc định.")
                 data['shd'] = ''
            if data and 'shd_type' not in data:
                 print(f"Phản hồi AI thiếu khóa 'shd_type', gán giá trị mặc định.")
                 data['shd_type'] = 'Khác'
            if data and 'cty' not in data:
                 print(f"Phản hồi AI thiếu khóa 'cty', gán giá trị mặc định.")
                 data['cty'] = 'Công ty không rõ'

            if data and 'ds' in data:
                data['ds'] = [item for item in data['ds'] if isinstance(item, dict)]
                if not data['ds']:
                     st.warning("⚠️ Danh sách thiết bị ('ds') trích xuất được trống hoặc không có mục hợp lệ.", icon="⚠️")
                     print("Danh sách thiết bị sau khi lọc rỗng.")
                     data = None

        except json.JSONDecodeError as e:
            st.error(f"❌ Lỗi khi giải mã JSON từ phản hồi AI: {e}. Phản hồi có thể không đúng định dạng JSON.", icon="❌")
            st.text_area("Phản hồi gốc từ AI:", a, height=200)
            print(f"Phản hồi AI gốc gây lỗi JSON: {a}")
            data = None
        except Exception as e:
            st.error(f"❌ Đã có lỗi không mong muốn khi xử lý dữ liệu từ AI: {e}", icon="❌")
            print(f"Dữ liệu nhận được trước lỗi: {data}")
            data = None

        # --- Logic gộp thiết bị và điền vào Word ---
        if data and 'ds' in data and data['ds']:
            st.info("✍️ Đang tạo file Word...", icon="⏳")
            try:
                # --- BƯỚC 1: Nhóm các thiết bị VÀ TÍNH TỔNG SỐ LƯỢNG GỘP ---
                grouped_devices = {}

                for item in data['ds']:
                    group_key_parts = []
                    group_key_parts.append(standardize_string(item.get('ttb', '')).strip())
                    group_key_parts.append(str(item.get('model', '')).strip())
                    group_key_parts.append(str(item.get('hang', '')).strip())
                    group_key_parts.append(str(item.get('nsx', '')).strip())
                    group_key_parts.append(str(item.get('dvt', '')).strip())
                    group_key_parts.append(str(item.get('pk', '')).strip())

                    group_key = tuple(group_key_parts)

                    current_sl_raw = item.get('sl', '0')
                    current_sl = 0
                    try:
                        cleaned_sl_str = re.sub(r'[^\d.]', '', str(current_sl_raw).strip())
                        current_sl = float(cleaned_sl_str) if cleaned_sl_str else 0
                    except (ValueError, TypeError):
                        print(f"Warning: Could not convert item quantity '{current_sl_raw}' to number during grouping. Defaulting to 0.")
                        current_sl = 0

                    current_seri = item.get('seri', [])
                    if not isinstance(current_seri, list):
                        current_seri = [current_seri]
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
                            'seri': cleaned_current_seri
                        }
                    else:
                        grouped_devices[group_key]['total_sl'] += current_sl
                        existing_seri_set = set(grouped_devices[group_key]['seri'])
                        new_seri_to_add = [s for s in cleaned_current_seri if s and s not in existing_seri_set]
                        grouped_devices[group_key]['seri'].extend(new_seri_to_add)


                # Bước 2: Chuyển đổi dictionary nhóm thành danh sách cuối cùng
                final_device_list = []
                for key, grouped_item in grouped_devices.items():
                    seri_string = ''
                    if grouped_item['seri']:
                         unique_seri = sorted(list(set(grouped_item['seri'])))
                         seri_string = 'Số seri: ' + ', '.join(unique_seri)
                    else:
                        seri_string = 'Số seri: Không có'

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

                # Bước 3: Điền dữ liệu vào bảng Word
                try:
                     document = Document('bbbg.docx')
                except Exception as e:
                     st.error(f"❌ Không tìm thấy hoặc không mở được file mẫu 'bbbg.docx'. Vui lòng đảm bảo file này nằm cùng thư mục với script.", icon="❌")
                     raise e

                font_name= 'Times New Roman'
                font_size=12

                print("\n--- Cấu trúc Paragraphs và Runs trong bbbg.docx ---")
                try:
                    for i, paragraph in enumerate(document.paragraphs):
                        print(f"Paragraph {i}: '{paragraph.text.strip()}'")
                        for j, run in enumerate(paragraph.runs):
                            print(f"  Run {j}: '{run.text}' (Length: {len(run.text)})")
                    print("-----------------------------------------------\n")
                except Exception as e:
                    print(f"Lỗi khi in cấu trúc Paragraphs và Runs: {e}")


                try:
                    table = document.tables[0]
                except IndexError:
                     st.error("❌ File mẫu 'bbbg.docx' không chứa bảng nào.", icon="❌")
                     raise IndexError

                if len(table.rows) > 1:
                    rows_to_remove_indices = range(len(table.rows) - 1, 0, -1)
                    for i in rows_to_remove_indices:
                        row = table.rows[i]
                        try:
                            tbl = row._tbl
                            tbl.getparent().remove(tbl)
                        except Exception as e:
                            print(f"Lỗi khi xóa hàng {i} trong bảng mẫu: {e}")

                count=0
                for item in final_device_list:
                    count += 1
                    ttb_text = str(item.get('ttb', '')).strip()
                    model_text = str(item.get('model', '')).strip()
                    hang_text = str(item.get('hang', '')).strip()
                    nsx_text = str(item.get('nsx', '')).strip()
                    dvt_text = str(item.get('dvt', '')).strip()
                    sl_text = str(int(item.get('sl', 0))).strip() if item.get('sl') is not None else ""
                    pk_text = str(item.get('pk', '')).strip()

                    device_info_text = f"{ttb_text}\n Hãng: {hang_text}\n NSX: {nsx_text}\n Model: {model_text}"
                    if pk_text:
                         device_info_text += f"\n{pk_text}"
                    else:
                         device_info_text += f"\nPhụ kiện: Không có"

                    new_device = [str(count),
                                  device_info_text,
                                  dvt_text,
                                  sl_text,
                                  item['seri_text']
                                 ]

                    row = table.add_row()
                    for i, cell_text in enumerate(new_device):
                        if i in (0, 2, 3):
                            ali = WD_ALIGN_PARAGRAPH.CENTER
                        else:
                            ali = WD_ALIGN_PARAGRAPH.LEFT
                        try:
                            cell = row.cells[i]
                            cell.text = str(cell_text)
                            for paragraph in cell.paragraphs:
                                paragraph.alignment = ali
                                for run in paragraph.runs:
                                    run.font.name = font_name
                                    run.font.size = Pt(font_size)
                        except IndexError:
                            st.warning(f"⚠️ Lỗi: Bảng trong file mẫu có ít hơn {len(new_device)} cột ({len(row.cells)}). Không thể điền dữ liệu đầy đủ cho hàng thiết bị thứ {count}.", icon="⚠️")
                            print(f"Lỗi: Hàng {count} có {len(row.cells)} ô, nhưng dữ liệu có {len(new_device)} mục.")
                            pass

                # --- Tìm và thay thế placeholder cho Số hợp đồng (ĐỊNH DẠNG THEO LOẠI) ---
                shd_value_raw = data.get('shd')
                shd_type_raw = data.get('shd_type')

                shd_value = str(shd_value_raw).strip() if shd_value_raw is not None else ''
                shd_type = str(shd_type_raw).strip() if shd_type_raw is not None else 'Khác'

                shd_value_to_replace = ''

                if shd_value:
                    shd_type_lower = shd_type.lower()

                    if 'hợp đồng' in shd_type_lower or 'hd' in shd_type_lower:
                        shd_value_to_replace = f"Dựa theo HĐ số: {shd_value}"
                    elif 'po' in shd_type_lower or 'đề nghị' in shd_type_lower or 'denghi' in shd_type_lower or 'mã đề nghị' in shd_type_lower:
                        shd_value_to_replace = f"Dựa theo PO: {shd_value}"
                    else:
                        shd_value_to_replace = f"Dựa theo số: {shd_value}"

                print(f"Value to replace placeholder with: '{shd_value_to_replace}' (Derived from value: '{shd_value}', type: '{shd_type}')")


                shd_placeholder_replaced = False
                shd_pattern = re.compile(re.escape("shd"), re.IGNORECASE)

                for paragraph in document.paragraphs:
                     if shd_pattern.search(paragraph.text):
                          for run in paragraph.runs:
                               original_run_text = run.text
                               new_run_text = shd_pattern.sub(shd_value_to_replace, original_run_text)

                               if new_run_text != original_run_text:
                                    run.text = new_run_text
                                    shd_placeholder_replaced = True

                if not shd_placeholder_replaced:
                     st.warning("⚠️ Không tìm thấy placeholder 'shd' (hoặc 'SHD',...) trong các đoạn văn của file mẫu. Số hợp đồng sẽ không được điền vào file Word.", icon="⚠️")
                     print("Không tìm thấy placeholder 'shd' (hoặc 'SHD',...).")

                # --- KẾT THÚC LOGIC THAY THẾ PLACEHOLDER (ĐỊNH DẠNG THEO LOẠI) ---

                # --- Tạo tên file đầu ra theo yêu cầu mới ---
                # Format: {Quantity}{DeviceName}-{Quantity}{DeviceName}_{ShortCompanyName}_{SHDValuePart}

                # 1. Chuỗi thông tin thiết bị (Số lượng + Tên thiết bị cho mỗi loại gộp)
                device_info_filename_parts = []
                for item in final_device_list:
                    quantity = int(item.get('sl', 0))
                    formatted_quantity = f"{quantity:02d}" if quantity >= 0 else "00"
                    device_name = str(item.get('ttb', '')).strip()

                    cleaned_device_name_part = re.sub(r'[\\/*?":<>|{}\[\]().,_]', '', device_name).strip()

                    if cleaned_device_name_part:
                         device_info_filename_parts.append(f"{formatted_quantity} {cleaned_device_name_part}")

                device_info_string_for_filename = "-".join(device_info_filename_parts)

                # 2. Lấy và rút gọn tên công ty (Bên giao)
                cty_name_raw = data.get('cty', 'UnknownCompany')
                cty_name_full = str(cty_name_raw).strip() if cty_name_raw is not None else 'UnknownCompany'
                cleaned_cty_name = shorten_company_name(cty_name_full)

                if not cleaned_cty_name:
                    cleaned_cty_name = re.sub(r'[\\/*?":<>|{}\[\]()]', '', cty_name_full).strip(" ,.-_&")


                # 3. Lấy giá trị SHD (chỉ phần số/mã trước dấu gạch ngang nếu có)
                shd_value_for_filename = shd_value

                shd_parts = shd_value_for_filename.split('-', 1)
                shd_cleaned_filename_part = shd_parts[0].strip() if shd_parts and shd_parts[0].strip() else ''

                shd_cleaned_filename_part = clean_filename(shd_cleaned_filename_part)


                # 4. Kết hợp các phần và làm sạch tên file lần cuối
                part1 = device_info_string_for_filename if device_info_string_for_filename else "ThietBi"
                part2 = cleaned_cty_name if cleaned_cty_name else "CongTy"
                part3 = shd_cleaned_filename_part if shd_cleaned_filename_part else "SoDinhDanh"

                raw_output_filename = f"{part1}_{part2}_{part3}"

                output_filename_final = clean_filename(raw_output_filename)

                if not output_filename_final.lower().endswith('.docx'):
                     output_filename = output_filename_final + '.docx'
                else:
                     output_filename = output_filename_final

                if not output_filename or output_filename.lower() == '.docx' or len(output_filename) < (len(".docx") + 3):
                    fallback_shd_part = shd_cleaned_filename_part if shd_cleaned_filename_part else "NoID"
                    fallback_cty_part = cleaned_cty_name if cleaned_cty_name else "CongTy"
                    output_filename = f"BienBanBanGiaoNoiBo_Fallback_{fallback_cty_part}_{fallback_shd_part}.docx"


                print(f"Generated output filename: {output_filename}")

                # --- KẾT THÚC TẠO TÊN FILE ĐẦU RA ---


                byte_io = BytesIO()
                document.save(byte_io)
                byte_io.seek(0)

                st.download_button(
                    label="✅ Tải xuống file Word Biên bản bàn giao nội bộ",
                    data=byte_io,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                st.success(f"🎉 Đã xử lý file PDF và tạo Biên bản bàn giao nội bộ: **{output_filename}**", icon="✅")

            except Exception as e:
                 st.error(f"❌ Đã có lỗi xảy ra trong quá trình tạo file Word: {e}", icon="❌")
                 print(f"Lỗi xử lý Word: {e}")

        elif data is not None:
             st.warning("⚠️ Không trích xuất được danh sách thiết bị nào từ file PDF.", icon="⚠️")
             print("Danh sách thiết bị 'ds' trống hoặc không hợp lệ.")

    except Exception as e:
        # Loại bỏ lỗi "No API_KEY" khỏi thông báo lỗi chung nếu nó đã được xử lý ở trên
        if "No API_KEY or ADC found" not in str(e):
             st.error(f"❌ Đã có lỗi xảy ra trong quá trình xử lý file: {e}", icon="❌")
        print(f"Lỗi chung khi xử lý file: {e}")

    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
                print(f"File tạm đã xóa: {temp_file_path}")
            except Exception as e:
                st.warning(f"⚠️ Lỗi khi xóa file tạm thời: {e}", icon="⚠️")
                print(f"Lỗi xóa file tạm: {e}")

elif google_api_key is not None:
     # Chỉ hiển thị thông báo chọn file nếu API key đã được cấu hình thành công
    st.info("⬆️ Vui lòng chọn một file PDF để bắt đầu.", icon="📄")

# else: # Trường hợp API key không cấu hình thành công, thông báo lỗi đã hiển thị ở trên
#    pass