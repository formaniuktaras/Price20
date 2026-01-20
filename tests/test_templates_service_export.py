import sys
import types
from datetime import datetime, timedelta
from pathlib import Path

import pytest

sys.path.append(str(Path(__file__).resolve().parents[1]))

jinja2_stub = types.ModuleType("jinja2")
jinja2_stub.Template = object
jinja2_stub.TemplateError = Exception
sys.modules.setdefault("jinja2", jinja2_stub)

import templates_service as ts


class DummyCell:
    def __init__(self, value):
        self.value = value
        self.alignment = None


class DummyAutoFilter:
    def __init__(self):
        self.ref = None


class DummyDimension:
    def __init__(self):
        self.width = None


class DummyColumnDimensions(dict):
    def __getitem__(self, key):
        if key not in self:
            self[key] = DummyDimension()
        return super().__getitem__(key)


class DummySheet:
    def __init__(self):
        self.rows = []
        self._cells = []
        self.title = ""
        self.freeze_panes = None
        self.auto_filter = DummyAutoFilter()
        self.column_dimensions = DummyColumnDimensions()

    @property
    def max_row(self):
        return len(self._cells)

    def append(self, row):
        values = tuple(row)
        self.rows.append(values)
        self._cells.append([DummyCell(value) for value in values])

    def iter_rows(self, min_row=1, max_row=None, max_col=None):
        start = max(min_row - 1, 0)
        end = max_row if max_row is not None else len(self._cells)
        for row in self._cells[start:end]:
            cells = list(row)
            if max_col is not None:
                if len(cells) < max_col:
                    cells.extend(DummyCell("") for _ in range(max_col - len(cells)))
                cells = cells[:max_col]
            yield tuple(cells)


class DummyWorkbook:
    def __init__(self, save_exc=None):
        self.active = DummySheet()
        self._save_exc = save_exc
        self.closed = False
        self.saved = None

    def save(self, filename):
        if self._save_exc:
            raise self._save_exc
        self.saved = filename
        Path(filename).touch()

    def close(self):
        self.closed = True


@pytest.fixture
def excel_context(monkeypatch):
    monkeypatch.setattr(ts, "OPENPYXL_AVAILABLE", True)
    monkeypatch.setattr(ts, "EXCEL_EXPORT_BLOCKED_MESSAGE", "")
    monkeypatch.setattr(ts, "OPENPYXL_IMPORT_ERROR_DETAIL", "")


def test_export_products_missing_dependency(monkeypatch, tmp_path):
    monkeypatch.setattr(ts, "OPENPYXL_AVAILABLE", False)
    monkeypatch.setattr(ts, "EXCEL_EXPORT_BLOCKED_MESSAGE", "")
    monkeypatch.setattr(ts, "OPENPYXL_IMPORT_ERROR_DETAIL", "")

    with pytest.raises(ts.ExportError) as excinfo:
        ts.export_products([], [], ts.EXCEL_FORMAT_LABEL, str(tmp_path))

    assert excinfo.value.code == ts.EXPORT_ERR_NO_OPENPYXL
    assert "Експорт у Excel" in excinfo.value.message


def test_export_products_permission_error(monkeypatch, tmp_path, excel_context):
    def make_workbook():
        return DummyWorkbook(PermissionError("denied"))

    monkeypatch.setattr(ts, "Workbook", make_workbook)

    with pytest.raises(ts.ExportError) as excinfo:
        ts.export_products([["value"]], ["col"], ts.EXCEL_FORMAT_LABEL, str(tmp_path))

    assert excinfo.value.code == ts.EXPORT_ERR_PERMISSION
    assert "доступ заборонено" in excinfo.value.message


def test_export_products_os_error(monkeypatch, tmp_path, excel_context):
    def make_workbook():
        return DummyWorkbook(OSError("disk full"))

    monkeypatch.setattr(ts, "Workbook", make_workbook)

    with pytest.raises(ts.ExportError) as excinfo:
        ts.export_products([["value"]], ["col"], ts.EXCEL_FORMAT_LABEL, str(tmp_path))

    assert excinfo.value.code == ts.EXPORT_ERR_OS_ERROR
    assert "disk full" in excinfo.value.message


def test_export_products_applies_formatting(monkeypatch, tmp_path, excel_context):
    created = {}

    def make_workbook():
        wb = DummyWorkbook()
        created["wb"] = wb
        return wb

    monkeypatch.setattr(ts, "Workbook", make_workbook)

    output = ts.export_products([["value", "long text"]], ["col", "desc"], ts.EXCEL_FORMAT_LABEL, str(tmp_path))

    sheet = created["wb"].active
    assert sheet.freeze_panes == "A2"
    assert sheet.auto_filter.ref and sheet.auto_filter.ref.startswith("A1:")
    width_value = sheet.column_dimensions.get("B")
    if hasattr(width_value, "width"):
        width = width_value.width
    else:
        width = width_value
    assert width is not None and width >= 10
    alignments = [
        cell.alignment for row in sheet.iter_rows(max_row=sheet.max_row, max_col=2) for cell in row
    ]
    assert any(getattr(al, "wrap_text", False) for al in alignments if al is not None)
    assert Path(output).exists()


def test_export_products_unknown_format():
    with pytest.raises(ts.ExportError) as excinfo:
        ts.export_products([], [], "custom", ".")

    assert excinfo.value.code == ts.EXPORT_ERR_UNKNOWN_FORMAT


def test_export_products_folder_preparation_failure(monkeypatch, tmp_path):
    target = tmp_path / "nested"

    def fake_exists(path):
        return False

    def fake_makedirs(path, exist_ok=True):
        raise OSError("cannot create directory")

    monkeypatch.setattr(ts.os.path, "exists", fake_exists)
    monkeypatch.setattr(ts.os, "makedirs", fake_makedirs)

    with pytest.raises(ts.ExportError) as excinfo:
        ts.export_products([], [], ts.CSV_FORMAT_LABEL, str(target))

    assert excinfo.value.code == ts.EXPORT_ERR_FOLDER_PREP
    assert "підготувати теку" in excinfo.value.message


def test_callable_datetime_behaves_like_datetime_operations():
    base = datetime(2024, 1, 15, 8, 30, 45)
    wrapped = ts._CallableDateTime(base)

    assert isinstance(wrapped, datetime)
    assert wrapped.year == base.year
    assert wrapped.strftime("%Y-%m-%d %H:%M:%S") == base.strftime("%Y-%m-%d %H:%M:%S")
    assert wrapped.replace(day=1).day == 1
    assert wrapped + timedelta(days=2) == base + timedelta(days=2)
    assert wrapped() == base


def test_callable_datetime_callable_result_supports_relativedelta():
    if ts._relativedelta is None:
        pytest.skip("relativedelta helper unavailable")

    base = datetime(2023, 1, 10, 12, 0, 0)
    wrapped = ts._CallableDateTime(base)
    result = wrapped().replace(day=1) + ts._relativedelta_helper(months=1, days=-1)

    expected = base.replace(day=1) + ts._relativedelta_helper(months=1, days=-1)
    assert result == expected
