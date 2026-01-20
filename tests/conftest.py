import os
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import pytest


@pytest.fixture(autouse=True)
def _force_data_dir(tmp_path, monkeypatch):
    data_dir = tmp_path / "data_dir"
    data_dir.mkdir(parents=True, exist_ok=True)
    monkeypatch.setenv("PRODGEN_DATA_DIR", str(data_dir))
    monkeypatch.setenv("PRICE16_DATA_DIR", str(data_dir))
    yield
