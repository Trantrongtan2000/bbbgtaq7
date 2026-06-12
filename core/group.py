"""Device grouping logic — merges identical devices by attributes."""

import json
from typing import List, Dict, Any
from core.models import Device, GroupedDevice
from utils.text import standardize_string

MAX_SERI_DISPLAY = 100


def group_devices(devices: List[Device]) -> List[GroupedDevice]:
    """Group identical devices by (ttb, model, hang, nsx, dvt, pk).

    Merges quantities and collects unique serial numbers.
    """
    grouped: Dict[tuple, Dict[str, Any]] = {}

    for device in devices:
        pk_key = (
            json.dumps(device.pk, ensure_ascii=False, sort_keys=True)
            if isinstance(device.pk, list)
            else str(device.pk).strip() if device.pk else ''
        )

        group_key = (
            standardize_string(device.ttb),
            device.model,
            device.hang,
            device.nsx,
            device.dvt,
            pk_key,
        )

        if group_key not in grouped:
            grouped[group_key] = {
                'ttb': device.ttb,
                'model': device.model,
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

    result = []
    for gd in grouped.values():
        unique_seri = sorted(gd['seri']) if gd['seri'] else []
        display_seri = unique_seri[:MAX_SERI_DISPLAY]
        seri_text = f"Số seri: {', '.join(display_seri)}"
        if len(unique_seri) > MAX_SERI_DISPLAY:
            seri_text += f" (và {len(unique_seri) - MAX_SERI_DISPLAY} seri khác)"

        result.append(GroupedDevice(
            ttb=gd['ttb'], model=gd['model'], hang=gd['hang'],
            nsx=gd['nsx'], dvt=gd['dvt'], sl=gd['total_sl'],
            pk=gd['pk_raw'], seri_text=seri_text,
        ))

    return result
