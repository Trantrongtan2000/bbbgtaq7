import pytest
from utils.text import standardize_string, clean_filename, shorten_company_name, convert_none_to_empty_string

def test_standardize_string():
 assert standardize_string('MAY TRUYEN DICH') == 'may truyen dich'
 assert standardize_string('Mindray') == 'mindray'
 assert standardize_string(' ECG ') == 'ecg'
 assert standardize_string(None) == ''

def test_clean_filename():
 assert clean_filename('Monitor 2.4G') == 'Monitor 2.4G'
 assert clean_filename('File/With:Bad?Chars*') == 'FileWithBadChars'

def test_shorten_company_name():
    assert shorten_company_name('CÔNG TY TNHH THƯƠNG MẠI MINH ĐỨC') == 'MINH ĐỨC'

def test_convert_none():
 data = {'a': None, 'b': 'val', 'c': [None, 'item']}
 assert convert_none_to_empty_string(data) == {'a': '', 'b': 'val', 'c': ['', 'item']}
