import pytest
from core.models import Device
from core.group import group_devices, _make_pk_key

def test_group_devices():
 d1 = Device(ttb='MAY THO', model=' SV300 ', hang='Mindray', sl=1, seri=['SN01'])
 d2 = Device(ttb='may tho', model='sv300', hang='MINDRAY', sl=2, seri=['SN02'])
 grouped = group_devices([d1, d2])
 assert len(grouped) == 1
 assert grouped[0].sl == 3
 assert 'SN01' in grouped[0].seri_text
 assert 'SN02' in grouped[0].seri_text

def test_pk_canonicalization():
 pk1 = ['day nguon', 'cap']
 pk2 = ['cap', 'day nguon']
 pk3 = [' Cap ', ' DAY NGUON ']
 assert _make_pk_key(pk1) == _make_pk_key(pk2)
 assert _make_pk_key(pk1) == _make_pk_key(pk3)

 d1 = Device(ttb='Bom tiem', pk=pk1, sl=1)
 d2 = Device(ttb='Bom tiem', pk=pk2, sl=2)
 grouped = group_devices([d1, d2])
 assert len(grouped) == 1
 assert grouped[0].sl == 3
