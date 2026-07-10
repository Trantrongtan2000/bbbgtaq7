"""Device grouping logic — merges identical devices by attributes."""

import json
from typing import List, Dict, Any
from core.models import Device, GroupedDevice
from utils.text import standardize_string

MAX_SERI_DISPLAY = 100


def _make_pk_key(pk: Any) -> str:
    """Create a hashable key from accessories list.

    Returns a string key for grouping. Handles None, empty lists,
    and serialization failures gracefully.
    """
    if pk is None:
        return "None"
    if isinstance(pk, list):
        if not pk:
            return "[]"
        try:
            return json.dumps(pk, ensure_ascii=False, sort_keys=True)
        except (TypeError, ValueError):
            return str(pk)
    return str(pk).strip() or "None"


def _make_group_key(device: Device) -> tuple:
    """Create a grouping key from device attributes.

    Uses getattr for safety in case Device structure changes.
    """
    return (
        standardize_string(getattr(device, 'ttb', '')),
        getattr(device, 'model', ''),
        getattr(device, 'ref', ''),
        getattr(device, 'hang', ''),
        getattr(device, 'nsx', ''),
        getattr(device, 'dvt', ''),
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
