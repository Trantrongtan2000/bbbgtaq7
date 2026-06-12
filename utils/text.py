"""Text normalization and string utility helpers."""

import re
from typing import Any


def standardize_string(text: Any) -> str:
    """Normalize Vietnamese diacritics to ASCII-safe equivalents."""
    if not isinstance(text, str):
        return str(text)

    replacements = [
        ('ÀÂẮẶẲẴ', 'AAAAA'), ('ÈÉẸẺẼ', 'EEEEE'), ('ỀẾỆỂỄ', 'EEEEE'),
        ('ÌÍỊỈĨ', 'IIIII'), ('ÒÓỌỎÕ', 'OOOOO'), ('ỒỐỘỔỖ', 'OOOOO'),
        ('ỜỚỢỞỠ', 'OOOOO'), ('ÙÚỤỦŨ', 'UUUUU'), ('ỪỨỰỬỮ', 'UUUUU'),
        ('ỲÝỴỶỸ', 'YYYYY'), ('Đ', 'D'),
    ]
    for src, dst in replacements:
        for s, d in zip(src, dst):
            text = text.replace(s, d)

    text = text.lower().replace('-', ' ').strip()
    return re.sub(r'\s+', ' ', text).strip()


def clean_filename(filename: str, max_len: int = 200) -> str:
    """Remove filesystem-illegal characters from filename."""
    chars_to_remove = r'[\\/*?":<>|.]'
    cleaned = re.sub(chars_to_remove, '', filename)
    return cleaned[:max_len] if len(cleaned) > max_len else cleaned


def shorten_company_name(company_name: str) -> str:
    """Shorten Vietnamese company names by removing legal entity suffixes."""
    if not isinstance(company_name, str):
        return str(company_name).strip()

    original = company_name.strip()
    name = original

    prefixes = [
        r"CÔNG TY TNHH MỘT THÀNH VIÊN", r"CÔNG TY TNHH MTV",
        r"CÔNG TY TNHH HAI THÀNH VIÊN TRỞ LÊN", r"CÔNG TY CỔ PHẦN",
        r"CÔNG TY TNHH", r"CÔNG TY", r"TNHH", r"CỔ PHẦN",
    ]
    suffixes = [
        r"MỘT THÀNH VIÊN", r"MTV", r"HAI THÀNH VIÊN TRỞ LÊN",
        r"CỔ PHẦN", r"TNHH",
    ]
    common_terms = [
        r"THƯƠNG MẠI VÀ DỊCH VỤ", r"DỊCH VỤ VÀ THƯƠNG MẠI",
        r"TM VÀ DV", r"DV VÀ TM", r"TM & DV", r"DV & TM",
        r"TM", r"DV", r"CÔNG NGHỆ", r"THƯƠNG MẠI", r"TRANG THIẾT BỊ",
        r"Y TẾ", r"XÂY DỰNG", r"ĐẦU TƯ", r"PHÁT TRIỂN", r"GIẢI PHÁP",
        r"KỸ THUẬT", r"SẢN XUẤT", r"NHẬP KHẨU", r"XUẤT NHẬP KHẨU",
        r"KINH DOANH", r"PHÂN PHỐI", r"VIỆT NAM"
    ]

    for p in prefixes + suffixes:
        name = re.sub(
            r'^\s*' + re.escape(p) + r'\s*|\s*' + re.escape(p) + r'\s*$',
            '', name, flags=re.IGNORECASE
        ).strip(" ,.-_&")

    for term in common_terms:
        name = re.sub(r'\b' + re.escape(term) + r'\b', '', name, flags=re.IGNORECASE).strip()
        name = re.sub(r'\s+', ' ', name).strip(" ,.-_&")

    return name if name else original


def convert_none_to_empty_string(obj: Any) -> Any:
    """Recursively convert None values to empty strings in dicts and lists."""
    if isinstance(obj, dict):
        return {k: convert_none_to_empty_string(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [convert_none_to_empty_string(elem) for elem in obj]
    return "" if obj is None else obj
