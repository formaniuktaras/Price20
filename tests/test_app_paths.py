import app_paths


def test_migrate_skips_when_target_exists(tmp_path):
    project_root = tmp_path / "project"
    project_root.mkdir()
    data_dir = app_paths.get_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)

    legacy = project_root / "catalog.db"
    legacy.write_text("legacy-data", encoding="utf-8")

    target = data_dir / "catalog.db"
    target.write_text("current-data", encoding="utf-8")

    app_paths.migrate_legacy_files(project_root)

    assert target.read_text(encoding="utf-8") == "current-data"
    assert not list(project_root.glob("catalog.db.bak_*"))
    assert not list(data_dir.glob("catalog.db.bak_*"))


def test_migrate_creates_backup_in_data_dir(tmp_path):
    project_root = tmp_path / "project"
    project_root.mkdir()
    data_dir = app_paths.get_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)

    legacy = project_root / "templates.json"
    legacy.write_text("legacy-templates", encoding="utf-8")

    app_paths.migrate_legacy_files(project_root)

    target = data_dir / "templates.json"
    assert target.exists()
    assert target.read_text(encoding="utf-8") == "legacy-templates"

    backups = list(data_dir.glob("templates.json.bak_*"))
    assert backups, "Expected a backup next to the migrated file"
    assert backups[0].read_text(encoding="utf-8") == "legacy-templates"
