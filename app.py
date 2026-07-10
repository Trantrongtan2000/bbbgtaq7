import streamlit as st
import os

from utils.logging_setup import get_logger
from utils.text import convert_none_to_empty_string
from config.api_keys import pool
from core.models import HandoverData
from core.group import group_devices
from core.filename import generate_filename
from core.extractor import extract_from_image, extract_text_from_pdf
from template.filler import fill_word_template

logger = get_logger('ui')

PROMPT_TEMPLATE = """
Hãy trích xuất thông tin từ Biên bản bàn giao và trả về JSON hợp lệ.

**Cấu trúc JSON bắt buộc:**
{
  "shd": "Số định danh (số hợp đồng, PO, v.v.)",
  "shd_type": "Loại: 'Hợp đồng', 'PO', 'Đề nghị', hoặc 'Khác'",
  "cty": "Tên công ty bên giao",
  "ds": [
    {
      "ttb": "Tên thiết bị",
      "model": "Model",
      "hang": "Hãng",
      "nsx": "Nước sản xuất",
      "dvt": "Đơn vị tính",
      "sl": Số lượng (số nguyên),
      "seri": ["danh sách số seri"] hoặc null,
      "pk": ["danh sách phụ kiện"] hoặc null
    }
  ]
}

**Quy tắc quan trọng:**
1. pk PHẢI là một ARRAY các chuỗi, ví dụ: ["Dây nguồn", "Cáp USB"]
2. pk KHÔNG được gộp thành một chuỗi dài
3. Nếu không có thông tin, trả về null
4. KHÔNG có Markdown code block, chỉ trả về JSON thuần
5. Đọc CHÍNH XÁC model/model number - cẩn thận với các số giống nhau (VD: 0/O, 1/I/l, 2/Z, 5/S, 6/G, 8/B)
6. Nếu không chắc chắn, ghi lại như trong tài liệu
"""


@st.cache_resource
def check_prerequisites() -> bool:
    if pool.size == 0:
        st.error("Không tìm thấy MISTRAL_API_KEY.")
        return False

    if not os.path.exists('bbbg.docx'):
        st.error("Thiếu file mẫu 'bbbg.docx'")
        return False

    return True


def main():
    st.set_page_config(
        page_title="Biên bản Bàn giao",
        page_icon="📄",
        layout="centered"
    )

    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Dark theme */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"],
    .main .block-container, section[data-testid="stSidebar"] > div {
        background-color: #0a0a0a !important;
    }

    /* Header bar */
    header[data-testid="stHeader"] {
        background-color: #0a0a0a !important;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #111111 !important;
    }

    /* Main content */
    .main .block-container {
        padding-top: 2rem !important;
        max-width: 640px !important;
    }

    /* Typography */
    h1 {
        font-family: 'Inter', -apple-system, sans-serif;
        font-weight: 700;
        letter-spacing: -0.03em;
        color: #f5f5f5 !important;
        font-size: 1.8rem !important;
    }

    .stMarkdown p {
        color: #a0a0a0;
        line-height: 1.6;
    }

    /* File uploader */
    .stFileUploader {
        border: 1.5px solid #333333;
        border-radius: 12px;
        padding: 1.5rem;
        background: #141414;
    }

    .stFileUploader:hover {
        border-color: #3b82f6;
    }

    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 2rem;
        transition: all 0.15s ease;
    }

    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #60a5fa 0%, #3b82f6 100%);
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
    }

    .stDownloadButton > button:active {
        transform: translateY(0);
    }

    /* Alerts */
    [data-testid="stSuccess"] {
        background-color: #0d2818 !important;
        border: 1px solid #166534 !important;
        border-radius: 10px !important;
        color: #4ade80 !important;
    }

    [data-testid="stError"] {
        background-color: #2d1215 !important;
        border: 1px solid #991b1b !important;
        border-radius: 10px !important;
        color: #f87171 !important;
    }

    [data-testid="stInfo"] {
        background-color: #141414 !important;
        border: 1px solid #333333 !important;
        border-radius: 10px !important;
        color: #a0a0a0 !important;
    }

    /* Spinner */
    [data-testid="stSpinner"] {
        color: #a0a0a0;
    }

    /* Divider */
    hr {
        border: none;
        border-top: 1px solid #222222;
    }

    /* File uploader label */
    .stFileUploader label {
        color: #a0a0a0 !important;
    }

    /* Input fields */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea {
        background-color: #141414 !important;
        border: 1px solid #333333 !important;
        color: #f5f5f5 !important;
    }

    /* Selectbox */
    .stSelectbox > div > div {
        background-color: #141414 !important;
        border: 1px solid #333333 !important;
        color: #f5f5f5 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    st.title("Biên bản Bàn giao")

    st.markdown("Tải lên file PDF hoặc ảnh từ Biên bản bàn giao để tạo file Word tự động.")

    if not check_prerequisites():
        st.stop()

    st.markdown(f"""
    <div style="
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 6px 12px;
        background: #0d2818;
        border: 1px solid #166534;
        border-radius: 6px;
        font-size: 0.85rem;
        color: #4ade80;
    ">
        <span style="width: 7px; height: 7px; background: #22c55e; border-radius: 50%; display: inline-block;"></span>
        Mistral OCR sẵn sàng &middot; {pool.size} key
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    uploaded_file = st.file_uploader(
        "Chọn file",
        type=["pdf", "jpg", "png"],
        help="Hỗ trợ định dạng PDF, PNG, JPG"
    )

    if uploaded_file:
        st.markdown(f"""
        <div style="
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 14px 16px;
            background: #141414;
            border: 1px solid #333333;
            border-radius: 10px;
            margin-bottom: 1rem;
        ">
            <div style="
                width: 40px;
                height: 40px;
                background: #1a1a1a;
                border-radius: 10px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 1.1rem;
            ">📄</div>
            <div>
                <div style="font-weight: 500; color: #f5f5f5; font-size: 0.9rem;">{uploaded_file.name}</div>
                <div style="font-size: 0.8rem; color: #666666;">Đang xử lý...</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Check PDF type (digital/searchable vs scanned) to inform the user
        is_digital_pdf = False
        num_pages = 0
        if uploaded_file.name.lower().endswith('.pdf'):
            try:
                pdf_text = extract_text_from_pdf(uploaded_file.getvalue())
                if pdf_text:
                    is_digital_pdf = True
            except Exception:
                pass
            try:
                import fitz
                doc = fitz.open(stream=uploaded_file.getvalue(), filetype="pdf")
                num_pages = len(doc)
            except Exception:
                pass

        if is_digital_pdf:
            st.info(f"Đã phát hiện văn bản trong PDF ({num_pages} trang). Bắt đầu trích xuất trực tiếp...")
        elif uploaded_file.name.lower().endswith('.pdf'):
            st.info(f"Không phát hiện văn bản trực tiếp. Đang chuyển đổi PDF ({num_pages} trang) sang ảnh và chạy Mistral OCR...")

        with st.spinner("Đang trích xuất dữ liệu..."):
            file_bytes = uploaded_file.getvalue()
            mime = 'application/pdf' if uploaded_file.name.lower().endswith('.pdf') else 'image/jpeg'
            data = extract_from_image(file_bytes, mime, PROMPT_TEMPLATE)

        if data and 'ds' in data:
            data = convert_none_to_empty_string(data)
            handover = HandoverData.from_dict(data)
            grouped = group_devices(handover.ds)

            filename = generate_filename(data, grouped)
            word_io = fill_word_template(data, grouped)

            st.markdown("""
            <div style="
                padding: 1rem 1.25rem;
                background: #0d2818;
                border: 1px solid #166534;
                border-radius: 10px;
                margin-bottom: 1rem;
            ">
                <div style="font-weight: 600; color: #4ade80; font-size: 0.9rem;">Trích xuất thành công</div>
                <div style="font-size: 0.85rem; color: #22c55e; margin-top: 2px;">File Word đã sẵn sàng tải xuống.</div>
            </div>
            """, unsafe_allow_html=True)

            st.download_button(
                "Tải file Word",
                word_io,
                filename,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.markdown("""
            <div style="
                padding: 1rem 1.25rem;
                background: #2d1215;
                border: 1px solid #991b1b;
                border-radius: 10px;
            ">
                <div style="font-weight: 600; color: #f87171; font-size: 0.9rem;">Không trích xuất được</div>
                <div style="font-size: 0.85rem; color: #ef4444; margin-top: 2px;">Vui lòng thử lại với file khác hoặc kiểm tra chất lượng ảnh.</div>
            </div>
            """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
