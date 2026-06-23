import streamlit as st
import os

from utils.logging_setup import get_logger
from utils.text import convert_none_to_empty_string
from config.api_keys import pool
from core.models import HandoverData
from core.group import group_devices
from core.filename import generate_filename
from core.extractor import extract_from_image
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
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

    .stApp {
        font-family: 'Inter', -apple-system, sans-serif;
    }

    h1 {
        font-family: 'Inter', sans-serif;
        font-weight: 700;
        letter-spacing: -0.03em;
        color: #1a1a2e;
        font-size: 2.2rem !important;
    }

    .stMarkdown p {
        color: #64748b;
        line-height: 1.6;
    }

    .stFileUploader {
        border: 1.5px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.5rem;
        transition: border-color 0.2s ease;
    }

    .stFileUploader:hover {
        border-color: #3b82f6;
    }

    .stDownloadButton > button {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 10px;
        padding: 0.75rem 2rem;
        letter-spacing: -0.01em;
        transition: all 0.2s ease;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }

    .stDownloadButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(30, 41, 59, 0.25);
    }

    .stDownloadButton > button:active {
        transform: translateY(0);
    }

    [data-testid="stSuccess"] {
        background-color: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 10px;
        color: #166534;
    }

    [data-testid="stError"] {
        background-color: #fef2f2;
        border: 1px solid #fecaca;
        border-radius: 10px;
        color: #991b1b;
    }

    [data-testid="stInfo"] {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        color: #475569;
    }

    .section-divider {
        border: none;
        border-top: 1px solid #f1f5f9;
        margin: 2rem 0;
    }
    </style>
    """, unsafe_allow_html=True)

    st.title("Biên bản Bàn giao")

    st.markdown("""
    <div style="max-width: 480px;">
        <p style="font-size: 1rem; color: #64748b; margin-bottom: 0;">
            Tải lên file PDF hoặc ảnh từ Biên bản bàn giao để tạo file Word tự động.
        </p>
    </div>
    """, unsafe_allow_html=True)

    if not check_prerequisites():
        st.stop()

    st.markdown(f"""
    <div style="
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 8px 14px;
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 8px;
        font-size: 0.875rem;
        color: #166534;
        font-weight: 500;
    ">
        <span style="
            width: 8px;
            height: 8px;
            background: #22c55e;
            border-radius: 50%;
            display: inline-block;
        "></span>
        Mistral OCR sẵn sàng &middot; {pool.size} key
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

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
            gap: 10px;
            padding: 12px 16px;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            margin-bottom: 1rem;
        ">
            <div style="
                width: 36px;
                height: 36px;
                background: #f1f5f9;
                border-radius: 8px;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 1rem;
            ">📄</div>
            <div>
                <div style="font-weight: 500; color: #1e293b; font-size: 0.9rem;">
                    {uploaded_file.name}
                </div>
                <div style="font-size: 0.8rem; color: #94a3b8; font-family: 'JetBrains Mono', monospace;">
                    Đang xử lý...
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

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
                background: linear-gradient(135deg, #f0fdf4 0%, #ecfdf5 100%);
                border: 1px solid #bbf7d0;
                border-radius: 12px;
                margin-bottom: 1rem;
            ">
                <div style="font-weight: 600; color: #166534; font-size: 0.95rem; margin-bottom: 4px;">
                    Trích xuất thành công
                </div>
                <div style="font-size: 0.85rem; color: #15803d;">
                    File Word đã sẵn sàng tải xuống.
                </div>
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
                background: #fef2f2;
                border: 1px solid #fecaca;
                border-radius: 12px;
            ">
                <div style="font-weight: 600; color: #991b1b; font-size: 0.95rem; margin-bottom: 4px;">
                    Không trích xuất được
                </div>
                <div style="font-size: 0.85rem; color: #b91c1c;">
                    Vui lòng thử lại với file khác hoặc kiểm tra chất lượng ảnh.
                </div>
            </div>
            """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
