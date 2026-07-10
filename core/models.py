"""Data models for handover document processing."""

from dataclasses import dataclass, field
from typing import List, Optional, Any, Dict


@dataclass
class Device:
    """Represents a single device in the handover document."""
    ttb: str = ""          # Tên thiết bị
    model: str = ""
    ref: str = ""          # Số REF
    hang: str = ""         # Hãng
    nsx: str = ""          # Nước sản xuất
    dvt: str = ""          # Đơn vị tính
    sl: float = 0          # Số lượng
    seri: List[str] = field(default_factory=list)
    pk: Optional[List[str]] = None  # Phụ kiện

    def to_dict(self) -> Dict[str, Any]:
        return {
            'ttb': self.ttb, 'model': self.model, 'ref': self.ref, 'hang': self.hang,
            'nsx': self.nsx, 'dvt': self.dvt, 'sl': self.sl,
            'seri': self.seri, 'pk': self.pk,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Device':
        sl_raw = data.get('sl', 0)
        try:
            sl = float(str(sl_raw).strip())
        except (ValueError, TypeError):
            sl = 0

        seri_raw = data.get('seri', []) or []
        if isinstance(seri_raw, str):
            seri = [seri_raw] if seri_raw else []
        elif isinstance(seri_raw, list):
            seri = [str(s).strip() for s in seri_raw if s and str(s).strip()]
        else:
            seri = []

        return cls(
            ttb=str(data.get('ttb', '')).strip(),
            model=str(data.get('model', '')).strip(),
            ref=str(data.get('ref', '')).strip(),
            hang=str(data.get('hang', '')).strip(),
            nsx=str(data.get('nsx', '')).strip(),
            dvt=str(data.get('dvt', '')).strip(),
            sl=sl,
            seri=seri,
            pk=data.get('pk'),
        )


@dataclass
class HandoverData:
    """Parsed handover document data."""
    shd: str = ""           # Số định danh
    shd_type: str = "Khác"  # Loại số
    cty: str = ""           # Tên công ty
    ds: List[Device] = field(default_factory=list)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'HandoverData':
        return cls(
            shd=str(data.get('shd', '')).strip(),
            shd_type=str(data.get('shd_type', 'Khác')).strip(),
            cty=str(data.get('cty', '')).strip(),
            ds=[Device.from_dict(d) for d in data.get('ds', []) if isinstance(d, dict)],
        )


@dataclass
class GroupedDevice:
    """Grouped device for Word template output."""
    ttb: str
    model: str
    ref: str
    hang: str
    nsx: str
    dvt: str
    sl: float
    pk: Any
    seri_text: str

    def to_dict(self) -> Dict[str, Any]:
        return {
            'ttb': self.ttb, 'model': self.model, 'ref': self.ref, 'hang': self.hang,
            'nsx': self.nsx, 'dvt': self.dvt, 'sl': self.sl,
            'pk': self.pk, 'seri_text': self.seri_text,
        }
