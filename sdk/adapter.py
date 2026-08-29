"""Mistral SDK adapter — OCR + chat completion with key rotation."""

import json
import base64
import time
from typing import Optional, Dict, Any, List
from mistralai import Mistral
from config.api_keys import pool
from utils.logging_setup import get_logger

logger = get_logger('sdk.adapter')


import re


def _parse_json_response(text: str) -> Optional[Dict[str, Any]]:
    """Parse JSON from model response supporting markdown code fences, raw JSON, and leading/trailing text."""
    if not text or not isinstance(text, str):
        return None

    cleaned = text.strip()

    # 1. Check for markdown code fences (```json ... ``` or ``` ...)
    fence_match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', cleaned, re.IGNORECASE)
    if fence_match:
        candidate = fence_match.group(1).strip()
        try:
            return json.loads(candidate)
        except Exception:
            cleaned = candidate

    # 2. Try direct json.loads on cleaned string
    try:
        return json.loads(cleaned)
    except Exception:
        pass

    # 3. Locate the first '{' and find the matching balanced '}'
    start_idx = cleaned.find('{')
    if start_idx != -1:
        depth = 0
        in_string = False
        escape = False
        for idx in range(start_idx, len(cleaned)):
            char = cleaned[idx]
            if escape:
                escape = False
                continue
            if char == '\\':
                escape = True
                continue
            if char == '"':
                in_string = not in_string
                continue
            if not in_string:
                if char == '{':
                    depth += 1
                elif char == '}':
                    depth -= 1
                    if depth == 0:
                        candidate = cleaned[start_idx:idx + 1]
                        try:
                            return json.loads(candidate)
                        except Exception:
                            break

    # 4. Fallback search for any outermost JSON object
    last_brace = cleaned.rfind('}')
    if start_idx != -1 and last_brace != -1 and last_brace > start_idx:
        try:
            return json.loads(cleaned[start_idx:last_brace + 1])
        except Exception as e:
            logger.warning(f"Failed to parse JSON substring: {e}")

    logger.warning("Could not extract valid JSON from response.")
    return None


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

