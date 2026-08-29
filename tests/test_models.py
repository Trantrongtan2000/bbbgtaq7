import pytest
from core.models import Device, HandoverData

def test_device_from_dict():
 raw = {'ttb': 'May do', 'sl': '2', 'seri': 'SN01, SN02; SN03'}
 dev = Device.from_dict(raw)
 assert dev.seri == ['SN01', 'SN02', 'SN03']
 assert dev.sl == 2.0

 raw2 = {'ttb': 'May do', 'sl': 1.5, 'seri': [' SN100 ', None, 'SN101 ']}
 dev2 = Device.from_dict(raw2)
 assert dev2.seri == ['SN100', 'SN101']
 assert dev2.sl == 1.5

def test_handover_data_from_dict():
 raw = {
 'shd': 'HD-12345',
 'shd_type': 'Hop dong',
 'cty': 'Cong ty ABC',
 'ds': [
 {'ttb': 'Device 1', 'sl': 1},
 {'ttb': 'Device 2', 'sl': 2.5}
 ]
 }
 hd = HandoverData.from_dict(raw)
 assert hd.shd == 'HD-12345'
 assert len(hd.ds) == 2
 assert hd.ds[1].sl == 2.5
