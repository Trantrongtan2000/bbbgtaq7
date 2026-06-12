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
    """Check API key availability and template file existence."""
    if pool.size == 0:
        st.error("❌ Không tìm thấy MISTRAL_API_KEY.", icon="❌")
        return False

    if not os.path.exists('bbbg.docx'):
        st.error("❌ Thiếu file mẫu 'bbbg.docx'", icon="❌")
        return False

    return True


def main():
    st.set_page_config(
        page_title="Chuyển đổi Bàn giao",
        page_icon="📋",
        layout="centered"
    )

    st.markdown("""
    <style>
    .stFileUploader { border: 2px dashed #004aad; padding: 1rem; border-radius: 0.5rem; }
    .stDownloadButton > button { background-color: #4CAF50; color: white; font-weight: bold; }
    h1 { color: #004aad; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

    st.title("📋 Chuyển đổi Biên bản Bàn giao")
    st.markdown("Tải lên file PDF hoặc ảnh từ Biên bản bàn giao để tạo file Word")

    if not check_prerequisites():
        st.stop()

    st.success(f"✅ Đã kết nối Mistral OCR! ({pool.size} API key)")

    uploaded_file = st.file_uploader(
        "📁 Chọn file (PDF/PNG/JPG)",
        type=["pdf", "jpg", "png"],
        help="Hỗ trợ định dạng PDF, PNG, JPG"
    )

    if uploaded_file:
        st.info(f"📥 Đang xử lý: **{uploaded_file.name}**", icon="⏳")

        file_bytes = uploaded_file.getvalue()
        mime = 'application/pdf' if uploaded_file.name.lower().endswith('.pdf') else 'image/jpeg'

        data = extract_from_image(file_bytes, mime, PROMPT_TEMPLATE)

        if data and 'ds' in data:
            data = convert_none_to_empty_string(data)
            handover = HandoverData.from_dict(data)
            grouped = group_devices(handover.ds)

            filename = generate_filename(data, grouped)
            word_io = fill_word_template(data, grouped)

            st.download_button(
                "⬇️ Tải xuống file Word",
                word_io,
                filename,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.success("✅ Hoàn thành! File Word đã sẵn sàng tải xuống.", icon="🎉")
        else:
            st.error("❌ Không trích xuất được dữ liệu từ file. Vui lòng thử file khác.", icon="❌")


if __name__ == "__main__":
    main()
