# Biên bản Bàn giao - Web App (Mistral OCR)

Web tạo biên bản bàn giao dựa trên biên bản giao hàng, sử dụng Mistral OCR để trích xuất dữ liệu.

## Cài đặt

1. **Cài đặt dependencies:**
```bash
pip install -r requirements.txt
```

2. **Cấu hình API Key:**
   - Tạo file `config.ini` trong thư mục project:
```ini
[API]
MISTRAL_API_KEY = YOUR_API_KEY_HERE
```
   - Hoặc đặt biến môi trường `MISTRAL_API_KEY`

3. **Chuẩn bị file mẫu:**
   - Đặt file `bbbg.docx` (Word template) trong thư mục project

## Chạy ứng dụng

```bash
streamlit run app.py
```

## Cách sử dụng

1. Mở ứng dụng Streamlit trên trình duyệt
2. Tải lên file PDF hoặc ảnh từ Biên bản bàn giao
3. Hệ thống sẽ trích xuất thông tin qua Mistral OCR và tạo file Word hoàn chỉnh
4. Tải xuống file Word được tạo tự động

## Tính năng

- Trích xuất dữ liệu từ PDF hoặc ảnh bằng Mistral OCR
- Xử lý 2 bước: OCR → trích xuất JSON bằng Mistral chat model
- Tự động nhận diện và phân tách thiết bị
- Xử lý danh sách phụ kiện (pk) dạng mảng
- Tạo tên file thông minh dựa trên nội dung
- Điền tự động vào template Word

## Cấu trúc project

```
bbbg-mistral/
├── app.py              # File chính của ứng dụng
├── config.ini          # File cấu hình (chứa API key)
├── bbbg.docx           # File template Word
├── requirements.txt    # Dependencies
├── core/               # Business logic
├── config/             # API key management
├── sdk/                # Mistral SDK adapter
├── template/           # Word template filling
├── utils/              # Text utilities
└── README.md           # File này
```

## Yêu cầu

- Python 3.8+
- Streamlit
- mistralai
- python-docx
