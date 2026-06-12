"""Smart filename generation for handover documents."""

import re
from typing import List, Dict, Any
from core.models import GroupedDevice
from utils.text import clean_filename, shorten_company_name

MAX_DEVICES_IN_FILENAME = 2


def generate_filename(data: Dict[str, Any], grouped_devices: List[GroupedDevice]) -> str:
    """Generate a descriptive filename like: 01_MayQuetTayCongNghe_ABC_12345.docx"""
    device_parts = []
    for item in grouped_devices[:MAX_DEVICES_IN_FILENAME]:
        sl = item.sl if hasattr(item, 'sl') else item.get('sl', 0)
        ttb = item.ttb if hasattr(item, 'ttb') else item.get('ttb', '')
        quantity = int(sl)
        formatted_quantity = f"{quantity:02d}"
        device_name = str(ttb).strip()
        cleaned_device_name = re.sub(r'[\\/*?":<>|{}\[\]().,_]', '', device_name).strip()
        if cleaned_device_name:
            device_parts.append(f"{formatted_quantity} {cleaned_device_name}")
    device_info_str = "-".join(device_parts) or "ThietBi"

    cty_name_full = str(data.get('cty', 'UnknownCompany')).strip()
    cleaned_cty_name = shorten_company_name(cty_name_full)
    if not cleaned_cty_name:
        cleaned_cty_name = re.sub(r'[\\/*?":<>|{}\[\]()]', '', cty_name_full).strip(" ,.-_&") or "CongTy"

    shd_value = str(data.get('shd', '')).strip()
    shd_main_part = shd_value.split('-', 1)[0].strip() or "SoDinhDanh"
    shd_cleaned = clean_filename(shd_main_part)

    raw_filename = f"{device_info_str}_{cleaned_cty_name}_{shd_cleaned}"
    final_filename_base = re.sub(r'\s+', '_', clean_filename(raw_filename)).strip('_')

    if not final_filename_base or len(final_filename_base) < 3:
        return f"BienBanBanGiao_{cleaned_cty_name}_{shd_cleaned}.docx"

    return final_filename_base + '.docx'
