"""Device grouping logic — merges identical devices by attributes."""

import json
from typing import List, Dict, Any
from core.models import Device, GroupedDevice
from utils.text import standardize_string

MAX_SERI_DISPLAY = 100


def _normalize_group_value(val: Any) -> str:
    """Normalize string value for grouping key: strip, lower, None-safe."""
    if val is None:
        return ""
    return str(val).strip().lower()


def _make_pk_key(pk: Any) -> str:
    """Create a canonicalized hashable key from accessories list.

    Strips each item, ignores casing and whitespace, removes empty items,
    and sorts the items before serialization.
    """
    if pk is None:
        return "None"
    if isinstance(pk, list):
        normalized_items = []
        for item in pk:
            if item is None:
                continue
            if isinstance(item, str):
                s = item.strip()
                if s:
                    normalized_items.append(s.lower())
            elif isinstance(item, (dict, list)):
                try:
                    normalized_items.append(json.dumps(item, ensure_ascii=False, sort_keys=True).lower())
                except Exception:
                    normalized_items.append(str(item).strip().lower())
            else:
                s = str(item).strip()
                if s:
                    normalized_items.append(s.lower())
        normalized_items.sort()
        try:
            return json.dumps(normalized_items, ensure_ascii=False)
        except (TypeError, ValueError):
            return str(normalized_items)
    return str(pk).strip().lower() or "None"


def _make_group_key(device: Device) -> tuple:
    """Create a grouping key from device attributes."""
    return (
        standardize_string(getattr(device, 'ttb', '')),
        _normalize_group_value(getattr(device, 'model', '')),
        _normalize_group_value(getattr(device, 'ref', '')),
        _normalize_group_value(getattr(device, 'hang', '')),
        _normalize_group_value(getattr(device, 'nsx', '')),
        _normalize_group_value(getattr(device, 'dvt', '')),
        _make_pk_key(getattr(device, 'pk', None)),
    )


def group_devices(devices: List[Device]) -> List[GroupedDevice]:
    """Group identical devices by (ttb, model, ref, hang, nsx, dvt, pk).

    Merges quantities and collects unique serial numbers.
    """
    grouped: Dict[tuple, Dict[str, Any]] = {}

    for device in devices:
        group_key = _make_group_key(device)

        if group_key not in grouped:
            grouped[group_key] = {
                'ttb': device.ttb,
                'model': device.model,
                'ref': device.ref,
                'hang': device.hang,
                'nsx': device.nsx,
                'dvt': device.dvt,
                'pk_raw': device.pk,
                'total_sl': device.sl,
                'seri': set(device.seri),
            }
        else:
            grouped[group_key]['total_sl'] += device.sl
            grouped[group_key]['seri'].update(device.seri)

    return [
        GroupedDevice(
            ttb=gd['ttb'], model=gd['model'], ref=gd['ref'], hang=gd['hang'],
            nsx=gd['nsx'], dvt=gd['dvt'], sl=gd['total_sl'],
            pk=gd['pk_raw'], seri_text=_format_seri(gd['seri']),
        )
        for gd in grouped.values()
    ]


def _format_seri(seri_set: set) -> str:
    """Format serial numbers for display.

    Returns formatted string with up to MAX_SERI_DISPLAY serial numbers.
    Appends remaining count if limit exceeded.
    """
    unique_seri = sorted(seri_set) if seri_set else []
    display_seri = unique_seri[:MAX_SERI_DISPLAY]
    text = f"Số seri: {', '.join(display_seri)}"
    if len(unique_seri) > MAX_SERI_DISPLAY:
        text += f" (và {len(unique_seri) - MAX_SERI_DISPLAY} seri khác)"
    return text
