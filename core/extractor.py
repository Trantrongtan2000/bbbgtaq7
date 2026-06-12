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


def extract_from_image(
    file_bytes: bytes,
    mime_type: str,
    prompt: str,
) -> Optional[Dict[str, Any]]:
    """Two-step extraction: Mistral OCR → chat model JSON parsing.

    Tries multiple API keys on quota errors.
    Returns parsed JSON dict or None.
    """
    last_error = None

    for attempt in range(pool.size):
        api_key = pool.get_current()
        if not api_key:
            break

        adapter = MistralAdapter(api_key)
        if not adapter.is_available:
            pool.rotate()
            continue

        try:
            ocr_text = adapter.ocr_document(file_bytes, mime_type)
            if not ocr_text:
                pool.rotate()
                continue

            text = adapter.chat_extract(ocr_text, prompt, SYSTEM_INSTRUCTION)
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
