import google.generativeai as genai
import streamlit as st
import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt  # For font size
from docx.enum.style import WD_STYLE_TYPE
import tempfile
from io import BytesIO

import json
import re

def clean_filename(filename):
    """Loại bỏ các ký tự đặc biệt khỏi tên file."""
    # Define the characters to remove
    chars_to_remove = (r'[\\/*?":<>|]')  # Raw string cho regex
    cleaned_name = re.sub(chars_to_remove, '', filename)
    return cleaned_name

st.title("Chuyển đổi Bàn giao cty thành bàn giao nội bộ")

GOOGLE_API_KEY = "AIzaSyDTyMSAO4W5-TUef4rtaAdm3J_4vU_LhT8"

genai.configure(api_key=GOOGLE_API_KEY)
model_name=['gemini-2.0-flash',"gemini-2.0-pro-exp-02-05","gemini-1.5-pro"]

file_name = st.file_uploader("Chọn file PDF", type="pdf")
print(file_name)
print(type(file_name))

if file_name is not None:  # Kiểm tra xem người dùng đã chọn file chưa
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_name.name.split('.')[-1]}") as temp_file:  # Create temp file
            temp_file.write(file_name.getvalue())  # Write the file content
            temp_file_path = temp_file.name  # Get the temporary file path
            print(temp_file_path)
        with st.spinner("Đang xử lý file..."):    
            file = file_name.name
            #print(file_name.read())
            sample_pdf = genai.upload_file(temp_file_path)
            model = genai.GenerativeModel(
                    model_name=model_name[1],
                    system_instruction=[
                        "Bạn là một nhà phân tích tài liệu kỹ thuật. Nhiệm vụ của bạn là:\
            - Trích xuất thông tin chi tiết quan trọng từ PDF với độ chính xác 98%\
            - Làm nổi bật mối quan hệ giữa dữ liệu trong bảng/biểu đồ\
            - Giữ nguyên cấu trúc tài liệu gốc trong đầu ra\
            - Đánh dấu sự mơ hồ trong các trang được quét",
            "Nhiệm vụ của bạn là hiểu dữ liệu trong tệp và trả lời các câu hỏi về dữ liệu đó.",
                    ],
            )
            prompt ='Dữ liệu đầu ra dạng json. Số hợp đồng (viết tắt là shd và chỉ 1 lần xuất hiện), danh sách thiết bị(viết tắt ds), Tên thiết bị (viết tắt ttb), model, hãng, nước sản xuất, đơn vị tính (viết tắt dvt),  số lượng, số seri (đầy đủ thông tin như tệp và một dòng), tên cty bên giao (viết tắt là cty và chỉ hiển thị 1 lần), phụ kiện (viết tắt là pk, chi tiết phụ kiện hoặc cấu hình kỹ thuật và có đẩy đủ số lượng, đơn vị tính, dữ liệu dạng chuỗi và xuống dòng ở từng mục và có chữ "Gồm:" ở dòng đầu tiên, nếu không có thì để trống). Yêu cầu như sau:\
            shd:\
            ttb:\
            model:\
            hang:\
            nsx:\
            dvt:\
            sl:\
            seri:\
            pk:\
            cty'
            response = model.generate_content([sample_pdf, prompt])
            for f in genai.list_files():
                print("  ", f.name)
                f.delete()
            a=response.text.replace("```", "")
            a=a.replace("json", "")
            try:
                data = json.loads(a)
            except json.JSONDecodeError as e:
                print(f"Error decoding JSON: {e}")
            document = Document('bbbg.docx')
            font_name= 'Times New Roman'
            font_size=12
            data1=[]
            table = document.tables[0]
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text)
                data1.append(row_data)

            count=0

            for item in data['ds']:
                count += 1
                new_device = [str(count),  # Count should be a string
                                f"{item['ttb']}\n Hãng: {item['hang']}\n NSX: {item['nsx']}\n Model: {item['model']}\n{item['pk']}",
                                str(item['dvt']) if item.get('dvt') is not None else "",  # dvt to string, handle missing key
                                str(item['sl']) if item.get('sl') is not None else "",  # sl to string, handle missing key
                                ('Số seri: ' + ', '.join(map(str, item['seri'])) if isinstance(item['seri'], list) else 'Số seri: ' + str(item['seri']) if item.get('seri') is not None else 'Số seri: ')]
                row = table.add_row()  # Add the row
                # Correct way to set cell content:
                for i, cell_text in enumerate(new_device):
                    if i in (0,2,3):
                        ali=WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        ali=WD_ALIGN_PARAGRAPH.LEFT
                    try:
                        cell = row.cells[i] # Get the individual cell
                        cell.text = cell_text # Set the cell text
                        for paragraph in cell.paragraphs:  # Important: Iterate through paragraphs!
                            paragraph.alignment = ali
                            for run in paragraph.runs:  # Apply to all runs in the paragraph
                                run.font.name = font_name
                                run.font.size = Pt(font_size)
                    except IndexError:
                        print(f"Error: Row has {len(row.cells)} cells, but new_device has {len(new_device)} items.")
                        break  # Exit the inner loop if there's an index error
            for paragraph in document.paragraphs:
                if "{shd}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{shd}", f"Theo HĐ số: {data['shd']}")
                    break
            shd=clean_filename(data['shd'])
            output_filename = f"{data['cty']}_{shd}.docx"
            document.save(output_filename)
            with open(output_filename, "rb") as file:
                    file_bytes = file.read()

            st.download_button(
                    label="Tải file word Biên bản bàn giao nội bộ.",
                    data=file_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", # Correct MIME type for .docx
                )

            st.success(f"Đã xử lý và sẵn sàng để tải xuống: {output_filename}")
    except Exception as e: # Bắt lỗi tổng quát hơn để hiển thị thông báo lỗi cho người dùng
        st.error(f"Đã có lỗi xảy ra trong quá trình xử lý: {e}")
    finally:
        try:
            if file_name is not None:
                os.remove(temp_file_path)  # Remove the temporary uploaded PDF
            for f in genai.list_files():
                f.delete()
        except Exception as e:
            st.warning(f"Lỗi khi xóa file tạm thời: {e}")
else:
    st.warning("Vui lòng chọn file PDF.")