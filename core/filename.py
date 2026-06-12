"""Smart filename generation for handover documents."""

import re
from typing import List, Dict, Any
from core.models import GroupedDevice
from utils.text import clean_filename, shorten_company_name

MAX_DEVICES_IN_FILENAME = 2


def _build_device_part(item: GroupedDevice) -> str:
    """Build filename part from a device: '01_TenThietBi'."""
    try:
        quantity = int(item.sl or 0)
    except (ValueError, TypeError):
        quantity = 0
    formatted_quantity = f"{quantity:02d}"
    device_name = str(item.ttb or '').strip()
    cleaned = re.sub(r'[\\/*?":<>|{}\[\]().,_]', '', device_name).strip()
    return f"{formatted_quantity} {cleaned}" if cleaned else ""


def _build_company_part(company_name: str) -> str:
    """Build shortened company name for filename."""
    shortened = shorten_company_name(company_name)
    if shortened:
        return re.sub(r'[\\/*?":<>|{}\[\]()]', '', shortened).strip(" ,.-_&") or "CongTy"
    return re.sub(r'[\\/*?":<>|{}\[\]()]', '', company_name).strip(" ,.-_&") or "CongTy"


def _build_shd_part(shd_value: str) -> str:
    """Build cleaned SHD identifier for filename."""
    if not shd_value:
        return "SoDinhDanh"
    main_part = shd_value.split('-', 1)[0].strip()
    return clean_filename(main_part) if main_part else "SoDinhDanh"


def generate_filename(data: Dict[str, Any], grouped_devices: List[GroupedDevice]) -> str:
    """Generate a descriptive filename like: 01_MayQuetTayCongNghe_ABC_12345.docx"""
    if not grouped_devices:
        return "BienBanBanGiao.docx"

    device_parts = [
        _build_device_part(item)
        for item in grouped_devices[:MAX_DEVICES_IN_FILENAME]
    ]
    device_info_str = "-".join(p for p in device_parts if p) or "ThietBi"

    company_name = str(data.get('cty', 'UnknownCompany')).strip()
    cleaned_company = _build_company_part(company_name)

    shd_value = str(data.get('shd', '')).strip()
    cleaned_shd = _build_shd_part(shd_value)

    raw_filename = f"{device_info_str}_{cleaned_company}_{cleaned_shd}"
    final_filename_base = re.sub(r'\s+', '_', clean_filename(raw_filename)).strip('_')

    if not final_filename_base or len(final_filename_base) < 5:
        return f"BienBanBanGiao_{cleaned_company}_{cleaned_shd}.docx"

    return final_filename_base + '.docx'
