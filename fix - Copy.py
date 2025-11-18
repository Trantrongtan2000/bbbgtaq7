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
from datetime import datetime

# System Instruction cho AI
SYSTEM_INSTRUCTION = (
    "Bạn là một nhà phân tích tài liệu kỹ thuật, chuyên trích xuất thông tin chi tiết từ 'Biên bản giao nhận - Nghiệm thu kiêm phiếu bảo hành' "
    "và các tài liệu tương tự. Nhiệm vụ của bạn là trích xuất các thông tin sau từ tệp PDF hoặc ảnh được cung cấp, đặc biệt là từ các bảng biểu, "
    "và **trả về DUY NHẤT dưới định dạng JSON hợp lệ**, không có bất kỳ văn bản giải thích, ký tự thừa, hoặc ký hiệu Markdown (như ```json) nào khác."
    "Sử dụng các viết tắt: shd (giá trị số định danh), shd_type (loại số định danh), cty, ds, ttb, model, hang, nsx, dvt, sl, seri, pk."
    "Lưu ý quan trọng: Nếu không tìm thấy Số seri hoặc Phụ kiện, hãy trả về giá trị là null cho các trường đó."
)

# --- Các hàm phụ trợ (Giữ nguyên) ---

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

def standardize_string(text):
    """Chuẩn hóa chuỗi tiếng Việt: loại bỏ dấu, chuyển lowercase, loại bỏ khoảng trắng thừa, dấu gạch ngang."""
    if not isinstance(text, str):
        return str(text)
    # Loại bỏ dấu
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
    # Các bước chuẩn hóa khác
    text = text.lower()
    text = re.sub(r'\s+', ' ', text).strip()
    text = text.replace('-', ' ').strip()
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def shorten_company_name(company_name):
    """Rút gọn tên công ty bằng cách loại bỏ các tiền tố và hậu tố phổ biến."""
    if not isinstance(company_name, str):
        return str(company_name).strip()

    cleaned_name = company_name.strip()
    
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

    # Loại bỏ tiền tố và hậu tố
    for p in prefixes + suffixes:
        cleaned_name = re.sub(r'^\s*' + re.escape(p) + r'\s*|' + r'\s*' + re.escape(p) + r'\s*$', '', cleaned_name, flags=re.IGNORECASE).strip(" ,.-_&")

    # Loại bỏ các từ phổ biến
    for term in common_terms:
        cleaned_name = re.sub(r'\b' + re.escape(term) + r'\b', '', cleaned_name, flags=re.IGNORECASE).strip()
        cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()

    cleaned_name = cleaned_name.strip(" ,.-_&")

    if not cleaned_name:
        words = company_name.strip().split()
        if words:
            # Fallback: lấy 3 từ cuối nếu tất cả bị loại bỏ
            fallback_name = " ".join(words[-3:]) 
            return fallback_name.strip(" ,.-_&")

        return company_name.strip()

    return cleaned_name
# --- Kết thúc các hàm phụ trợ ---


def process_and_generate_word_doc(data, raw_ai_response_text):
    """
    Processes the extracted data from AI and generates the Word document.
    """
    try:
        # Convert all None values to empty strings recursively
        data = convert_none_to_empty_string(data)
        print("Dữ liệu JSON nhận được (sau khi xử lý None):", json.dumps(data, indent=2, ensure_ascii=False))

        extracted_shd = data.get('shd')
        extracted_shd_type = data.get('shd_type')
        print(f"Extracted shd value from AI: '{extracted_shd}' (Type: '{extracted_shd_type}')")

        # Validation và gán giá trị mặc định
        if 'ds' not in data or not isinstance(data.get('ds'), list):
            st.error("❌ Phản hồi từ AI không chứa danh sách thiết bị hợp lệ ('ds'). Vui lòng thử lại với file khác hoặc kiểm tra nội dung file.", icon="❌")
            print(f"Phản hồi AI thiếu khóa 'ds' hoặc 'ds' không phải list: {data}")
            return False 
        
        if data and 'shd' not in data: data['shd'] = ''
        if data and 'shd_type' not in data: data['shd_type'] = 'Khác'
        if data and 'cty' not in data: data['cty'] = 'Công ty không rõ'

        if data and 'ds' in data:
            data['ds'] = [item for item in data['ds'] if isinstance(item, dict)]
            if not data['ds']:
                st.warning("⚠️ Danh sách thiết bị ('ds') trích xuất được trống hoặc không có mục hợp lệ.", icon="⚠️")
                print("Danh sách thiết bị sau khi lọc rỗng.")
                return False 

    except Exception as e:
        st.error(f"❌ Đã có lỗi không mong muốn khi xử lý dữ liệu từ AI: {e}", icon="❌")
        print(f"Lỗi xử lý dữ liệu: {e}. Dữ liệu nhận được trước lỗi: {data}")
        return False

    # --- Logic gộp thiết bị và điền vào Word ---
    if data and 'ds' in data and data['ds']:
        st.info("✍️ Đang tạo file Word...", icon="⏳")
        try:
            # --- BƯỚC 1: Nhóm các thiết bị VÀ TÍNH TỔNG SỐ LƯỢNG GỘP ---
            grouped_devices = {}

            for item in data['ds']:
                # Dùng các trường đã chuẩn hóa (lowercase, không dấu) làm khóa nhóm
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
                current_sl = 0
                try:
                    # Loại bỏ ký tự không phải số hoặc dấu chấm
                    cleaned_sl_str = re.sub(r'[^\d.]', '', str(current_sl_raw).strip())
                    current_sl = float(cleaned_sl_str) if cleaned_sl_str else 0
                except (ValueError, TypeError):
                    print(f"Warning: Could not convert item quantity '{current_sl_raw}' to number during grouping. Defaulting to 0.")
                    current_sl = 0

                # Xử lý Seri
                current_seri = item.get('seri', [])
                # Do đã chuyển None thành "" ở bước đầu, ta kiểm tra giá trị:
                if isinstance(current_seri, str) and not current_seri:
                    current_seri = [] # Coi chuỗi rỗng là danh sách rỗng
                elif not isinstance(current_seri, list):
                    current_seri = [current_seri] if current_seri else []
                # Làm sạch và loại bỏ chuỗi rỗng
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
                    # Gộp seri, tránh trùng lặp
                    existing_seri_set = set(grouped_devices[group_key]['seri'])
                    new_seri_to_add = [s for s in cleaned_current_seri if s and s not in existing_seri_set]
                    grouped_devices[group_key]['seri'].extend(new_seri_to_add)


            # Bước 2: Chuyển đổi dictionary nhóm thành danh sách cuối cùng
            final_device_list = []
            for key, grouped_item in grouped_devices.items():
                seri_string = ''
                if grouped_item['seri']:
                    unique_seri = sorted(list(set(grouped_item['seri'])))
                    # Giới hạn số lượng seri hiển thị trên một dòng
                    display_seri = unique_seri
                    if len(unique_seri) > 100:
                        display_seri = unique_seri[:100]
                        seri_string = 'Số seri: ' + ', '.join(display_seri) + f" (và {len(unique_seri) - 100} seri khác)"
                    else:
                        seri_string = 'Số seri: ' + ', '.join(unique_seri)
                else:
                    # Yêu cầu: "Số seri: Không có" thành ""
                    seri_string = '' 

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

            try:
                table = document.tables[0]
            except IndexError:
                st.error("❌ File mẫu 'bbbg.docx' không chứa bảng nào.", icon="❌")
                raise IndexError

            # Xóa các hàng dữ liệu mẫu (trừ hàng tiêu đề đầu tiên)
            if len(table.rows) > 1:
                # Xóa ngược từ dưới lên
                for i in range(len(table.rows) - 1, 0, -1):
                    row = table.rows[i]
                    try:
                        tbl = row._tbl
                        tbl.getparent().remove(tbl)
                    except Exception as e:
                        print(f"Lỗi khi xóa hàng {i} trong bảng mẫu: {e}")

            # Thêm các hàng mới
            count=0
            for item in final_device_list:
                count += 1
                ttb_text = str(item.get('ttb', '')).strip()
                model_text = str(item.get('model', '')).strip()
                hang_text = str(item.get('hang', '')).strip()
                nsx_text = str(item.get('nsx', '')).strip()
                dvt_text = str(item.get('dvt', '')).strip()
                # Chuyển số lượng thành chuỗi số nguyên (ví dụ: 1.0 -> 1)
                sl_text = str(int(item.get('sl', 0))).strip() if item.get('sl') is not None else ""
                pk_text = str(item.get('pk', '')).strip()

                device_info_text = f"{ttb_text}\n- Model: {model_text}\n- Hãng: {hang_text}\n- NSX: {nsx_text}"
                
                # --- Xử lý Phụ kiện (pk) ---
                pk_output_text = ""
                if pk_text:
                    # remove "Cấu hình bao gồm:" and similar phrases, and leading `-`
                    pk_text = re.sub(r'(cấu hình bao gồm|bao gồm|chi tiết cấu hình):','', pk_text, flags=re.IGNORECASE).strip()
                    pk_text = pk_text.replace('–', '-').strip() # Chuẩn hóa gạch ngang
                    accessories = pk_text.split('\n')
                    # Indent accessories, lọc bỏ dòng trống
                    accessories = [f"  + {acc.strip().lstrip('-').lstrip('•').strip()}" for acc in accessories if acc.strip()]
                    
                    if accessories:
                        pk_output_text = "\n- Phụ kiện:\n" + "\n".join(accessories)
                
                # Nối pk_output_text (nếu không có phụ kiện hợp lệ, nó là "")
                device_info_text += pk_output_text
                # --- Kết thúc xử lý Phụ kiện ---

                new_device = [str(count),
                              device_info_text,
                              dvt_text,
                              sl_text,
                              item['seri_text'] # Giá trị là chuỗi rỗng hoặc danh sách seri (có tiền tố)
                             ]

                row = table.add_row()
                for i, cell_text in enumerate(new_device):
                    # Căn giữa cột STT, ĐVT, SL. Các cột khác căn trái
                    ali = WD_ALIGN_PARAGRAPH.CENTER if i in (0, 2, 3) else WD_ALIGN_PARAGRAPH.LEFT
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

            # --- Thay thế ngày tháng năm thực tế vào dòng Tp.HCM, ngày ... ---
            now = datetime.now()
            current_day = str(now.day)
            current_month = str(now.month)
            current_year = str(now.year)
            for paragraph in document.paragraphs:
                if "Tp.HCM" in paragraph.text and ("day" in paragraph.text or "month" in paragraph.text or "year" in paragraph.text):
                    new_text = paragraph.text
                    new_text = new_text.replace("day", current_day).replace("month", current_month).replace("year", current_year)
                    paragraph.text = new_text

            # --- Tìm và thay thế placeholder cho Số hợp đồng (ĐỊNH DẠNG THEO LOẠI) ---
            shd_value_raw = data.get('shd')
            shd_type_raw = data.get('shd_type')

            shd_value = str(shd_value_raw).strip() if shd_value_raw is not None else ''
            shd_type = str(shd_type_raw).strip() if shd_type_raw is not None else 'Khác'

            shd_value_to_replace = ''

            if shd_value:
                shd_type_lower = standardize_string(shd_type)

                if 'hop dong' in shd_type_lower or 'hd' in shd_type_lower:
                    shd_value_to_replace = f"Dựa theo HĐ số: {shd_value}"
                elif 'po' in shd_type_lower or 'de nghi' in shd_type_lower or 'denghi' in shd_type_lower or 'ma de nghi' in shd_type_lower:
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
            
            # 1. Chuỗi thông tin thiết bị (Số lượng + Tên thiết bị cho mỗi loại gộp)
            device_info_filename_parts = []
            for item in final_device_list:
                quantity = int(item.get('sl', 0))
                formatted_quantity = f"{quantity:02d}" if quantity >= 0 else "00"
                device_name = str(item.get('ttb', '')).strip()

                cleaned_device_name_part = re.sub(r'[\\/*?":<>|{}\[\]().,_]', '', device_name).strip()

                if cleaned_device_name_part:
                    device_info_filename_parts.append(f"{formatted_quantity} {cleaned_device_name_part}")

            device_info_string_for_filename = "-".join(device_info_filename_parts[:2]) # Giới hạn 2 thiết bị đầu cho gọn

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

            # Tên file cuối cùng: {DeviceName(s)}_{ShortCompanyName}_{SHDValuePart}.docx
            raw_output_filename = f"{part1}_{part2}_{part3}"
            
            # Xử lý làm sạch lần cuối, thay khoảng trắng bằng gạch dưới, và giới hạn độ dài
            output_filename_final = re.sub(r'\s+', '_', clean_filename(raw_output_filename))
            output_filename_final = output_filename_final.strip('_')

            output_filename = output_filename_final + '.docx'

            # Fallback nếu tên file quá ngắn hoặc chỉ có đuôi
            if not output_filename_final or len(output_filename_final) < 3:
                fallback_shd_part = shd_cleaned_filename_part if shd_cleaned_filename_part else "NoID"
                fallback_cty_part = cleaned_cty_name if cleaned_cty_name else "CongTy"
                output_filename = f"BienBanBanGiaoNoiBo_Fallback_{fallback_cty_part}_{fallback_shd_part}.docx"


            print(f"Generated output filename: {output_filename}")

            # --- KẾT THÚC TẠO TÊN FILE ĐẦU RA ---

            # Tạo file download
            byte_io = BytesIO()
            document.save(byte_io)
            byte_io.seek(0)

            st.download_button(
                label="✅ Tải xuống file Word Biên bản bàn giao nội bộ",
                data=byte_io,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

            st.success(f"🎉 Đã xử lý file và tạo Biên bản bàn giao nội bộ: **{output_filename}**", icon="✅")
            return True # Indicate success

        except Exception as e:
            st.error(f"❌ Đã có lỗi xảy ra trong quá trình tạo file Word: {e}", icon="❌")
            print(f"Lỗi xử lý Word: {e}")
            return False 

    elif data is not None:
        st.warning("⚠️ Không trích xuất được danh sách thiết bị nào từ file PDF.", icon="⚠️")
        print("Danh sách thiết bị 'ds' trống hoặc không hợp lệ.")
        return False 

    return False # Default return if data is None or other issues

# --- Cấu hình giao diện và CSS (Giữ nguyên) ---
st.set_page_config(page_title="Chuyển đổi Bàn giao", layout="centered")

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
# --- Kết thúc Cấu hình giao diện và CSS ---

# ----------------------------------------------------------------------
## Cấu hình Google API Key
# ----------------------------------------------------------------------
config = configparser.ConfigParser()
config_file_path = 'config.ini'
google_api_key = None 

# Biến cờ để kiểm tra cấu hình API thành công
is_api_configured = False 

if os.path.exists(config_file_path):
    config.read(config_file_path)
    try:
        # Lấy API Key
        google_api_key = config['API']['GEMINI_API_KEY'] 
        
        # Cấu hình API Key bằng genai.configure (Phương pháp tương thích nhất)
        genai.configure(api_key=google_api_key)
        
        is_api_configured = True 
        print("Đã đọc API Key từ config.ini và cấu hình genai.")

    except KeyError:
        st.error(f"❌ Lỗi cấu hình: File '{config_file_path}' không có section [API] hoặc key GEMINI_API_KEY. Vui lòng kiểm tra lại file config.ini.", icon="❌")
        google_api_key = None 
    except Exception as e:
        st.error(f"❌ Lỗi khi đọc file cấu hình '{config_file_path}': {e}. Vui lòng kiểm tra định dạng file config.ini.", icon="❌")
        google_api_key = None 
else:
    st.error(f"❌ Lỗi cấu hình: Không tìm thấy file cấu hình '{config_file_path}'. Vui lòng tạo file này với section [API] và key GEMINI_API_KEY.", icon="❌")
    google_api_key = None 
# --- Kết thúc cấu hình API Key ---
# ----------------------------------------------------------------------

st.title("Công cụ Chuyển đổi Biên bản Bàn giao")
st.subheader("Tải lên Biên bản bàn giao gốc (PDF hoặc Ảnh)")
file_name = st.file_uploader("Chọn file Biên bản bàn giao công ty (PDF hoặc Ảnh)", type=["pdf", "jpg", "jpeg", "png"], label_visibility="collapsed", key="file_uploader")

st.markdown("ℹ️ **Lưu ý:** File mẫu Word (`bbbg.docx`) phải nằm cùng thư mục với script.")


file_bytes = None
file_mime_type = None

# Danh sách các model để thử nghiệm theo thứ tự ưu tiên
MODEL_PRIORITY_LIST = [
    "gemini-2.5-pro", 
    "gemini-2.5-flash", 
    "gemini-2.5-flash-lite",
    "gemini-2.0-flash" 
]

# Chỉ tiếp tục xử lý nếu có file được tải lên VÀ API Key đã được cấu hình thành công
if file_name is not None and is_api_configured:
    try:
        st.info(f"📥 Đang tải lên và xử lý file: **{file_name.name}**", icon="⏳")
        
        # 1. Đọc file dưới dạng bytes và xác định MIME type
        file_bytes = file_name.getvalue()
        file_extension = file_name.name.split('.')[-1].lower()
        
        if file_extension == 'pdf':
            file_mime_type = 'application/pdf'
        elif file_extension in ['jpg', 'jpeg']:
            file_mime_type = 'image/jpeg'
        elif file_extension == 'png':
            file_mime_type = 'image/png'
        else:
            raise ValueError("Định dạng file không được hỗ trợ để truyền trực tiếp.")

        # 2. Tạo đối tượng Part để truyền trực tiếp
        uploaded_file_part = {
            'mime_type': file_mime_type,
            'data': file_bytes
        }

        # 3. Chuẩn bị Prompt
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
        # 4. Vòng lặp thử nghiệm các model
        data = None
        raw_ai_response = ""
        model_used = None

        for model_name in MODEL_PRIORITY_LIST:
            try:
                with st.spinner(f"✨ Đang trích xuất dữ liệu từ file bằng model: **{model_name}** (Ưu tiên: {MODEL_PRIORITY_LIST.index(model_name) + 1})..."):
                    
                    # Khởi tạo model 
                    model = genai.GenerativeModel(
                        model_name=model_name,
                        system_instruction=SYSTEM_INSTRUCTION 
                    )
                    
                    # Gọi generate_content (Tương thích cao nhất, không dùng config/mime_type)
                    response = model.generate_content(
                        contents=[uploaded_file_part, prompt_content]
                    )
                    
                    raw_ai_response = response.text
                    print(f"Raw AI response from {model_name}: {raw_ai_response}")
                    
                    # Cố gắng làm sạch và tải JSON
                    a = raw_ai_response.strip()
                    if a.startswith("```json"):
                        a = a[len("```json"):].strip()
                    if a.endswith("```"):
                        a = a[:-len("```")].strip()
                        
                    data = json.loads(a)
                    model_used = model_name
                    break # Thành công, thoát vòng lặp
                    
            except Exception as e:
                # Báo lỗi và thử model tiếp theo
                st.warning(f"⚠️ Model {model_name} gặp lỗi hoặc không trả về JSON hợp lệ: {e}. Đang thử model tiếp theo...", icon="⚠️")
                print(f"Model {model_name} failed: {e}")
                data = None
                raw_ai_response = response.text if 'response' in locals() and response else ""


        # 5. Xử lý phản hồi cuối cùng
        if data is None:
            st.error("❌ Tất cả các model đã thử đều không thể trích xuất dữ liệu JSON hợp lệ. Vui lòng kiểm tra lại file đầu vào hoặc prompt.", icon="❌")
            if raw_ai_response:
                st.text_area("Phản hồi gốc cuối cùng từ AI (gây lỗi):", raw_ai_response, height=200)
        else:
            process_and_generate_word_doc(data, raw_ai_response)

    except Exception as e:
        if "No API_KEY" not in str(e):
            st.error(f"❌ Đã có lỗi xảy ra trong quá trình xử lý file: {e}", icon="❌")
        print(f"Lỗi chung khi xử lý file: {e}")

    finally:
        # Không cần xử lý file tạm
        pass

elif is_api_configured:
    # Chỉ hiển thị thông báo chọn file nếu API key đã được cấu hình thành công
    st.info("⬆️ Vui lòng chọn một file PDF/Ảnh để bắt đầu.", icon="📄")