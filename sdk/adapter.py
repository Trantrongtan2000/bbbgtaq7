"""Mistral SDK adapter — OCR + chat completion with key rotation."""

import json
import base64
import time
from typing import Optional, Dict, Any, List
from mistralai.client import Mistral
from config.api_keys import pool
from utils.logging_setup import get_logger

logger = get_logger('sdk.adapter')


def _parse_json_response(text: str) -> Optional[Dict[str, Any]]:
    """Strip markdown code fences and parse JSON."""
    text = text.strip()
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    return json.loads(text.strip())


class MistralAdapter:
    """Mistral OCR + Chat adapter with key rotation."""

    def __init__(self, api_key: str):
        self._api_key = api_key
        self._client = Mistral(api_key=api_key)

    @property
    def is_available(self) -> bool:
        return self._client is not None

    def ocr_document(self, file_bytes: bytes, mime_type: str) -> Optional[str]:
        """OCR a document/image using Mistral OCR. Returns markdown text."""
        try:
            b64_data = base64.b64encode(file_bytes).decode('utf-8')
            data_url = f"data:{mime_type};base64,{b64_data}"

            if mime_type == 'application/pdf':
                doc = {"type": "document_url", "document_url": data_url}
            else:
                doc = {"type": "image_url", "image_url": data_url}

            ocr_response = self._client.ocr.process(
                model="mistral-ocr-latest",
                document=doc,
            )

            pages = ocr_response.pages if ocr_response.pages else []
            parts = [p.markdown for p in pages if p.markdown]
            result = "\n\n".join(parts)

            if result.strip():
                logger.info(f"OCR succeeded: {len(result)} chars")
                return result
            logger.warning("OCR returned empty content")
            return None

        except Exception as e:
            logger.error(f"OCR failed: {type(e).__name__}: {e}")
            raise

    def chat_extract(
        self,
        ocr_text: str,
        prompt: str,
        system_instruction: str,
    ) -> Optional[str]:
        """Send OCR text to Mistral chat for structured extraction."""
        try:
            response = self._client.chat.complete(
                model="mistral-large-latest",
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": f"{prompt}\n\n---\nNội dung OCR:\n{ocr_text}"},
                ],
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            logger.error(f"Chat extraction failed: {type(e).__name__}: {e}")
            raise


def extract_from_image(
    file_bytes: bytes,
    mime_type: str,
    prompt: str,
) -> Optional[Dict[str, Any]]:
    """Two-step extraction: OCR → JSON parsing via Mistral.

    Tries multiple API keys on failure.
    """
    system_instruction = (
        "Bạn là một nhà phân tích tài liệu kỹ thuật. Nhiệm vụ của bạn là trích xuất thông tin từ 'Biên bản bàn giao' "
        "vào định dạng JSON. "
        "QUAN TRỌNG: Trường 'pk' (Phụ kiện) phải là một danh sách (Array) các chuỗi, không được gộp thành 1 chuỗi dài. "
        "Nếu không có thông tin, trả về null. Không thêm Markdown (```json)."
    )

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
            # Step 1: OCR
            ocr_text = adapter.ocr_document(file_bytes, mime_type)
            if not ocr_text:
                pool.rotate()
                continue

            # Step 2: Chat extraction
            text = adapter.chat_extract(ocr_text, prompt, system_instruction)
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
