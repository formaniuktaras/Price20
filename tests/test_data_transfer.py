import data_transfer as dt
import templates_service as ts


class DummySheet:
    def __init__(self, workbook, title="Sheet", rows=None):
        self._workbook = workbook
        self._title = title
        self.rows = [] if rows is None else list(rows)
        self._workbook._register_sheet(self)

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, value):
        if self._title in self._workbook.sheets:
            del self._workbook.sheets[self._title]
        self._title = value
        self._workbook.sheets[value] = self

    def append(self, row):
        self.rows.append(tuple(row))

    def iter_rows(self, min_row=1, values_only=False):
        start = max(min_row - 1, 0)
        for row in self.rows[start:]:
            yield row


class DummyWorkbook:
    def __init__(self):
        self.sheets = {}
        self.active = DummySheet(self)
        self.saved = None
        self.closed = False

    def _register_sheet(self, sheet):
        self.sheets[sheet.title] = sheet

    def create_sheet(self, title):
        sheet = DummySheet(self, title=title)
        return sheet

    def save(self, filename):
        self.saved = filename

    def close(self):
        self.closed = True


def test_export_all_data_to_excel(monkeypatch, tmp_path):
    sample_catalog = {
        "categories": [{"id": 1, "name": "Cat", "created_at": "2024-01-01"}],
        "brands": [
            {"id": 2, "category_id": 1, "name": "Brand", "created_at": "2024-01-02"}
        ],
        "models": [
            {"id": 3, "brand_id": 2, "name": "Model", "created_at": "2024-01-03"}
        ],
        "specs": [
            {"id": 4, "model_id": 3, "key": "Key", "value": "Value"}
        ],
    }

    created = {}

    def fake_workbook():
        wb = DummyWorkbook()
        created["wb"] = wb
        return wb

    monkeypatch.setattr(dt, "OPENPYXL_AVAILABLE", True)
    monkeypatch.setattr(dt, "EXCEL_EXPORT_BLOCKED_MESSAGE", "")
    monkeypatch.setattr(dt, "Workbook", fake_workbook)
    monkeypatch.setattr(dt, "export_catalog_dump", lambda: sample_catalog)

    filename = tmp_path / "backup.xlsx"
    templates = {
        "title_template": "{{ title }}",
        "tags_template": "{{ tags }}",
        "descriptions": {
            "Cat": {
                "прозора": {
                    "default": "Опис укр",
                    "languages": {"en": "Description en"},
                }
            }
        },
        "template_languages": [{"code": "uk", "label": "Українська"}],
        "film_types": [{"name": "прозора", "enabled": True}],
    }
    title_tags = {
        "default": {
            "title_template": {"default": "{{ title }}", "languages": {"en": "Title EN"}},
            "tags_template": {"default": "{{ tags }}"},
        },
        "by_film": {
            "прозора": {
                "title_template": {"default": "Film title"},
                "tags_template": {"default": "Film tags", "languages": {"uk": "Плівка теги"}},
            }
        },
        "by_category": {
            "Cat": {
                "default": {
                    "title_template": {"default": "Cat title"},
                }
            }
        },
    }
    export_fields = [
        {"field": "Назва", "template": "{{ title }}", "enabled": True, "languages": ["uk", "en"]}
    ]

    dt.export_all_data_to_excel(str(filename), templates, title_tags, export_fields)

    wb = created["wb"]
    assert wb.saved == str(filename)
    assert "Категорії" in wb.sheets
    assert wb.sheets["Категорії"].rows[1] == (1, "Cat", "2024-01-01")
    assert wb.sheets["Параметри"].rows[1][0] == "language"
    assert wb.sheets["Параметри"].rows[2][0] == "film_type"
    assert wb.sheets["Експортні поля"].rows[1][4] == "uk, en"

    templates_sheet = wb.sheets["Шаблони"].rows
    assert templates_sheet[0] == (
        "Група",
        "Сценарій",
        "Поле",
        "Категорія",
        "Тип плівки",
        "Мова",
        "Значення",
    )
    assert ("templates", "global", "title_template", "", "", "", "{{ title }}") in templates_sheet
    assert (
        "templates",
        "description_language",
        "description",
        "Cat",
        "прозора",
        "en",
        "Description en",
    ) in templates_sheet
    assert (
        "title_tags",
        "film_language",
        "tags_template",
        "",
        "прозора",
        "uk",
        "Плівка теги",
    ) in templates_sheet
    assert (
        "title_tags",
        "category_default",
        "title_template",
        "Cat",
        "",
        "",
        "Cat title",
    ) in templates_sheet


class LoadedSheet:
    def __init__(self, rows):
        self.rows = [tuple(row) for row in rows]

    def iter_rows(self, min_row=1, values_only=True):
        start = max(min_row - 1, 0)
        for row in self.rows[start:]:
            yield row


class LoadedWorkbook:
    def __init__(self, sheets):
        self.sheetnames = list(sheets.keys())
        self._sheets = {name: LoadedSheet(rows) for name, rows in sheets.items()}
        self.closed = False

    def __getitem__(self, item):
        return self._sheets[item]

    def close(self):
        self.closed = True


def test_import_all_data_from_excel(monkeypatch):
    sheets = {
        "Категорії": [("ID", "Назва", "Створено"), (1, "Cat", "2024-01-01")],
        "Бренди": [
            ("ID", "Категорія ID", "Назва", "Створено"),
            (2, 1, "Brand", "2024-01-02"),
        ],
        "Моделі": [
            ("ID", "Бренд ID", "Назва", "Створено"),
            (3, 2, "Model", "2024-01-03"),
        ],
        "Характеристики": [
            ("ID", "Модель ID", "Ключ", "Значення"),
            (4, 3, "Key", "Value"),
        ],
        "Параметри": [
            ("Тип", "Код", "Назва", "Увімкнено"),
            ("language", "uk", "Українська", ""),
            ("film_type", "", "прозора", "1"),
        ],
        "Шаблони": [
            (
                "Група",
                "Сценарій",
                "Поле",
                "Категорія",
                "Тип плівки",
                "Мова",
                "Значення",
            ),
            ("templates", "global", "title_template", "", "", "", "{{ title }}"),
            ("templates", "global", "tags_template", "", "", "", "{{ tags }}"),
            ("templates", "description_default", "description", "Cat", "прозора", "", "Опис"),
            (
                "templates",
                "description_language",
                "description",
                "Cat",
                "прозора",
                "en",
                "Description EN",
            ),
            (
                "title_tags",
                "default",
                "title_template",
                "",
                "",
                "",
                "{{ title }}",
            ),
            (
                "title_tags",
                "default_language",
                "title_template",
                "",
                "",
                "en",
                "Title EN",
            ),
            (
                "title_tags",
                "film",
                "tags_template",
                "",
                "прозора",
                "",
                "Film tags",
            ),
            (
                "title_tags",
                "film_language",
                "tags_template",
                "",
                "прозора",
                "uk",
                "Плівка теги",
            ),
            (
                "title_tags",
                "category_default",
                "title_template",
                "Cat",
                "",
                "",
                "Cat title",
            ),
        ],
        "Експортні поля": [
            ("Позиція", "Поле", "Шаблон", "Увімкнено", "Мови"),
            (1, "Назва", "{{ title }}", "1", "uk"),
        ],
    }

    workbook = LoadedWorkbook(sheets)

    monkeypatch.setattr(dt, "OPENPYXL_AVAILABLE", True)
    monkeypatch.setattr(dt, "EXCEL_EXPORT_BLOCKED_MESSAGE", "")
    monkeypatch.setattr(dt, "_load_workbook_file", lambda filename: workbook)
    monkeypatch.setattr(dt, "Workbook", object())
    monkeypatch.setattr(dt.os.path, "exists", lambda path: True)

    captured = {}

    def fake_import_catalog_dump(payload):
        captured["catalog"] = payload

    monkeypatch.setattr(dt, "import_catalog_dump", fake_import_catalog_dump)

    saved_templates = {}
    saved_title_tags = {}
    saved_export_fields = {}

    monkeypatch.setattr(ts, "save_templates", lambda data: saved_templates.update(data=data))
    monkeypatch.setattr(ts, "save_title_tags_templates", lambda data: saved_title_tags.update(data=data))
    monkeypatch.setattr(ts, "save_export_fields", lambda data: saved_export_fields.update(data=data))

    monkeypatch.setattr(ts, "load_templates", lambda: {"loaded": True})
    monkeypatch.setattr(ts, "load_title_tags_templates", lambda templates: {"titles": True})
    monkeypatch.setattr(ts, "load_export_fields", lambda: [{"field": "Назва"}])

    result = dt.import_all_data_from_excel("dummy.xlsx")

    assert captured["catalog"]["categories"][0]["name"] == "Cat"
    assert saved_templates["data"]["title_template"] == "{{ title }}"
    assert (
        saved_templates["data"]["descriptions"]["Cat"]["прозора"]["languages"]["en"]
        == "Description EN"
    )
    assert saved_templates["data"]["template_languages"][0]["code"] == "uk"
    assert saved_templates["data"]["film_types"][0]["name"] == "прозора"
    assert saved_title_tags["data"]["default"]["title_template"]["languages"]["en"] == "Title EN"
    assert (
        saved_title_tags["data"]["by_film"]["прозора"]["tags_template"]["languages"]["uk"]
        == "Плівка теги"
    )
    assert (
        saved_title_tags["data"]["by_category"]["Cat"]["default"]["title_template"]["default"]
        == "Cat title"
    )
    assert saved_export_fields["data"][0]["field"] == "Назва"
    assert result == {
        "templates": {"loaded": True},
        "title_tags_templates": {"titles": True},
        "export_fields": [{"field": "Назва"}],
    }

