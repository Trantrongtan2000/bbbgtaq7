"""AI extraction — Mistral OCR + chat with retry + key rotation."""

import json
import time
from typing import Optional, Dict, Any
from config.api_keys import pool
from sdk.adapter import MistralAdapter, _parse_json_response
from utils.logging_setup import get_logger

logger = get_logger('core.extractor')


SYSTEM_INSTRUCTION = (
    "Bạn là một nhà phân tích tài liệu kỹ thuật. Nhiệm vụ của bạn là trích xuất thông tin từ 'Biên bản bàn giao' "
    "vào định dạng JSON. "
    "QUAN TRỌNG: Trường 'pk' (Phụ kiện) phải là một danh sách (Array) các chuỗi, không được gộp thành 1 chuỗi dài. "
    "Nếu không có thông tin, trả về null. Không thêm Markdown (```json)."
)


def extract_text_from_pdf(file_bytes: bytes) -> str:
    """Extract text from PDF using PyMuPDF (fitz) with fallback to pypdf."""
    # 1. Try PyMuPDF (fitz)
    try:
        import fitz
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        text = ""
        for page in doc:
            page_text = page.get_text()
            if page_text:
                text += page_text + "\n"
        text = text.strip()
        if len(text) > 50:
            logger.info(f"PyMuPDF successfully extracted {len(text)} chars from PDF.")
            return text
    except Exception as e:
        logger.warning(f"PyMuPDF text extraction failed: {e}")

    # 2. Fallback to pypdf
    try:
        import io
        import pypdf
        reader = pypdf.PdfReader(io.BytesIO(file_bytes))
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
        text = text.strip()
        if len(text) > 50:
            logger.info(f"pypdf successfully extracted {len(text)} chars from PDF.")
            return text
    except Exception as e:
        logger.warning(f"pypdf text extraction failed: {e}")

    return ""


def convert_pdf_to_images(file_bytes: bytes) -> list[bytes]:
    """Convert PDF pages to a list of PNG image bytes using PyMuPDF."""
    import fitz
    images = []
    try:
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            img_data = pix.tobytes("png")
            images.append(img_data)
        logger.info(f"Successfully converted PDF to {len(images)} images.")
    except Exception as e:
        logger.error(f"Error converting PDF to images: {e}")
    return images


def extract_from_image(
    file_bytes: bytes,
    mime_type: str,
    prompt: str,
) -> Optional[Dict[str, Any]]:
    """Two-step extraction: PDF text extraction (or PDF page-by-page image OCR) -> chat model JSON parsing.

    Tries multiple API keys on quota errors.
    Returns parsed JSON dict or None.
    """
    last_error = None

    # Step 0: Try direct PDF text extraction if it's a PDF
    ocr_text = ""
    pdf_images = []
    if mime_type == 'application/pdf':
        try:
            ocr_text = extract_text_from_pdf(file_bytes)
        except Exception as e:
            logger.warning(f"Direct PDF text extraction failed: {e}")

        # If no direct text was found, convert PDF to images for page-by-page OCR
        if not ocr_text:
            try:
                pdf_images = convert_pdf_to_images(file_bytes)
            except Exception as e:
                logger.warning(f"Failed to convert PDF to images: {e}")

    for attempt in range(pool.size):
        api_key = pool.get_current()
        if not api_key:
            break

        adapter = MistralAdapter(api_key)
        if not adapter.is_available:
            pool.rotate()
            continue

        try:
            # Step 1: Get text (either already extracted, via page-by-page OCR, or single image OCR)
            current_ocr_text = ocr_text
            if not current_ocr_text:
                if mime_type == 'application/pdf' and pdf_images:
                    # Run OCR on each page image and combine
                    ocr_parts = []
                    for i, img_bytes in enumerate(pdf_images):
                        logger.info(f"Running OCR on PDF page {i+1}/{len(pdf_images)}")
                        page_text = adapter.ocr_document(img_bytes, 'image/png')
                        if page_text:
                            ocr_parts.append(page_text)
                        else:
                            raise ValueError(f"OCR returned empty content for page {i+1}")
                    current_ocr_text = "\n\n".join(ocr_parts)
                else:
                    # Normal single image OCR
                    current_ocr_text = adapter.ocr_document(file_bytes, mime_type)

            if not current_ocr_text:
                pool.rotate()
                continue

            # Step 2: Chat extraction
            text = adapter.chat_extract(current_ocr_text, prompt, SYSTEM_INSTRUCTION)
            if not text:
                pool.rotate()
                continue

            data = _parse_json_response(text)
            if data:
                logger.info(f"Successfully extracted data with key index {pool._index}")
                return data

        except Exception as e:
            last_error = str(e)
            logger.warning(f"Attempt {attempt + 1} failed: {type(e).__name__}: {last_error}")

            quota_errors = [
                "API_KEY", "UNAUTHORIZED", "INVALID", "quota", "limit",
                "429", "RESOURCE_EXHAUSTED", "rate_limit", "PERMISSION_DENIED",
            ]
            if any(x in last_error.upper() for x in quota_errors):
                logger.warning("Quota/key error, rotating to next key...")
                pool.rotate()
                time.sleep(0.5)
                continue

            pool.rotate()
            continue

    if last_error:
        logger.error(f"All keys failed: {last_error}")
    return None


