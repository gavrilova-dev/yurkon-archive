from app.window import make_protocol_filename


def test_protocol_filename_replaces_slash_and_removes_quotes():
    assert make_protocol_filename('ПР-145/26', 'ООО «Ромашка»') == 'ПР-145-26 ООО Ромашка'


def test_protocol_filename_strips_surrounding_spaces_and_empty_parts():
    assert make_protocol_filename('  145/26  ', '  «Организация»  ') == '145-26 Организация'
    assert make_protocol_filename('', 'ООО Ромашка') == 'ООО Ромашка'
