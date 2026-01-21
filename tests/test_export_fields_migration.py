import sys
import types
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1]))

jinja2_stub = types.ModuleType("jinja2")
jinja2_stub.Template = object
jinja2_stub.TemplateError = Exception
sys.modules.setdefault("jinja2", jinja2_stub)

import templates_service as ts


def test_migrate_export_field_template_for_code():
    fields = [
        {
            "field": "Код_товару",
            "template": "{{ spec('Код_товару') }}",
            "enabled": True,
        }
    ]

    updated, changed = ts._migrate_export_fields_templates(fields)

    assert changed is True
    assert updated[0]["field"] == "Код_товару"
    assert updated[0]["template"] == "{{ clean_id(spec('Код_товару')) }}"


def test_migrate_export_field_template_keeps_custom_template():
    fields = [
        {
            "field": "Код_товару",
            "template": "{{ spec('Код_товару') }}-{{ brand }}",
            "enabled": True,
        }
    ]

    updated, changed = ts._migrate_export_fields_templates(fields)

    assert changed is False
    assert updated[0]["field"] == "Код_товару"
    assert updated[0]["template"] == "{{ spec('Код_товару') }}-{{ brand }}"
