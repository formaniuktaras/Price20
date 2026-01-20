from specs_io import format_specs_for_clipboard, parse_specs_payload


def test_parse_specs_payload_supports_multiple_separators():
    raw = "Вага; 1.2 кг\nКолір:\tчорний\nАкумулятор=4000 мАг"
    assert parse_specs_payload(raw) == [
        ("Вага", "1.2 кг"),
        ("Колір", "чорний"),
        ("Акумулятор", "4000 мАг"),
    ]


def test_parse_specs_payload_ignores_headers():
    raw = "Назва параметра;Значення\nВага;1 кг\nКолір;Чорний"
    assert parse_specs_payload(raw) == [("Вага", "1 кг"), ("Колір", "Чорний")]


def test_parse_specs_payload_strips_bom_characters():
    raw = "\ufeffНазва параметра;Значення\n\ufeffВага;1 кг"
    assert parse_specs_payload(raw) == [("Вага", "1 кг")]


def test_parse_specs_payload_skips_comments_and_empty_lines():
    raw = "# comment\n\nМатеріал корпусу: метал\n\n"
    assert parse_specs_payload(raw) == [("Матеріал корпусу", "метал")]


def test_format_specs_for_clipboard_builds_tab_delimited_string():
    specs = [("Вага", "1"), ("Колір", "")]
    assert format_specs_for_clipboard(specs) == "Вага\t1\nКолір\t"


def test_format_specs_for_clipboard_trims_values():
    specs = [("  Назва  ", "  Значення "), ("", None)]
    assert format_specs_for_clipboard(specs) == "Назва\tЗначення\n\t"
