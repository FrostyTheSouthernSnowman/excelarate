from excellent._sheet import SpreadSheet


def test_spreadsheet_init():
    sheet = SpreadSheet("test")
    assert True


def test_spreadsheet_format_dict_basic():
    sheet = SpreadSheet("test")
    assert sheet._format_dict({"a": "b", "c": 2, 3: "d", "e": True}) == {
        "a": "b", "c": "2", "3": "d", "e": "true"}


def test_spreadsheet_format_dict_nested_dicts():
    sheet = SpreadSheet("test")
    assert sheet._format_dict({"a": {"b": "c", "d": 2, 2: False}}) == {
        "a_b": "c", "a_d": '2', "a_2": "false"}


def test_spreadsheet_format_dict_lists():
    sheet = SpreadSheet("test")
    assert sheet._format_dict({"a": ["b", "c", "d"]}) == {
        "a": "['b', 'c', 'd']"}
