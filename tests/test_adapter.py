import pytest
from sdk.adapter import _parse_json_response

def test_parse_raw_json():
    text = '{"shd": "123", "ds": []}'
    res = _parse_json_response(text)
    assert res == {'shd': '123', 'ds': []}

def test_parse_fenced_json():
    text = '```json\n{"shd": "123", "ds": []}\n```'
    res = _parse_json_response(text)
    assert res == {'shd': '123', 'ds': []}

def test_parse_prose():
    text = 'Here is the extracted result:\n{"shd": "123", "ds": [{"ttb": "A"}]}\nHope this helps!'
    res = _parse_json_response(text)
    assert res == {'shd': '123', 'ds': [{'ttb': 'A'}]}

def test_parse_malformed():
    assert _parse_json_response('Not a json string') is None

