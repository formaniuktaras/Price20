from datetime import datetime

import pytest

import database


def test_export_catalog_dump_returns_entities():
    database.init_db()

    database.add_category("Категорія A")
    cat_id = next(cid for cid, name in database.get_categories() if name == "Категорія A")
    database.add_brand(cat_id, "Бренд A")
    brand_id = next(bid for bid, name in database.get_brands(cat_id) if name == "Бренд A")
    database.add_model(brand_id, "Модель A")
    model_id = next(mid for mid, name in database.get_models(brand_id) if name == "Модель A")
    database.insert_spec(model_id, "Ключ", "Значення")

    dump = database.export_catalog_dump()

    assert dump["categories"] and dump["brands"] and dump["models"] and dump["specs"]
    assert any(entry["name"] == "Категорія A" for entry in dump["categories"])
    assert any(entry["category_id"] == cat_id and entry["name"] == "Бренд A" for entry in dump["brands"])
    assert any(entry["brand_id"] == brand_id and entry["name"] == "Модель A" for entry in dump["models"])
    assert any(entry["model_id"] == model_id and entry["key"] == "Ключ" for entry in dump["specs"])


def test_import_catalog_dump_replaces_data():
    database.init_db()

    database.add_category("OldCat")
    old_cat = next(cid for cid, name in database.get_categories() if name == "OldCat")
    database.add_brand(old_cat, "OldBrand")
    old_brand = next(bid for bid, name in database.get_brands(old_cat) if name == "OldBrand")
    database.add_model(old_brand, "OldModel")

    payload = {
        "categories": [{"id": 1, "name": "NewCat", "created_at": "2024-01-01 00:00:00"}],
        "brands": [
            {
                "id": 1,
                "category_id": 1,
                "name": "NewBrand",
                "created_at": "2024-01-01 00:00:00",
            }
        ],
        "models": [
            {
                "id": 1,
                "brand_id": 1,
                "name": "NewModel",
                "created_at": "2024-01-01 00:00:00",
            }
        ],
        "specs": [
            {"id": 1, "model_id": 1, "key": "Новий ключ", "value": "123"},
        ],
    }

    database.import_catalog_dump(payload)

    categories = database.get_categories()
    assert categories == [(1, "NewCat")]
    brands = database.get_brands(1)
    assert brands == [(1, "NewBrand")]
    models = database.get_models(1)
    assert models == [(1, "NewModel")]
    specs = database.get_specs(1)
    assert [(sid, key, value) for sid, key, value in specs] == [(1, "Новий ключ", "123")]


def _prepare_model() -> int:
    database.init_db()

    database.add_category("Тестова")
    categories = database.get_categories()
    cat_id = next(cid for cid, name in categories if name == "Тестова")

    database.add_brand(cat_id, "BrandX")
    brands = database.get_brands(cat_id)
    brand_id = next(bid for bid, name in brands if name == "BrandX")

    database.add_model(brand_id, "ModelX")
    models = database.get_models(brand_id)
    model_id = next(mid for mid, name in models if name == "ModelX")
    return model_id


def test_insert_spec_returns_row_id():
    model_id = _prepare_model()

    spec_id = database.insert_spec(model_id, "Параметр", "Значення")
    assert isinstance(spec_id, int) and spec_id > 0

    specs = database.get_specs(model_id)
    assert len(specs) == 1
    stored_id, key, value = specs[0]
    assert stored_id == spec_id
    assert key == "Параметр"
    assert value == "Значення"


def test_replace_specs_resets_previous_values():
    model_id = _prepare_model()
    first_id = database.insert_spec(model_id, "Old", "123")
    assert first_id

    database.replace_specs(
        model_id,
        [
            ("  Key A  ", "  Value A  "),
            ("Key B", "Value B"),
            ("", "should be ignored"),
        ],
    )

    specs = database.get_specs(model_id)
    assert [(key, value) for _sid, key, value in specs] == [
        ("Key A", "  Value A  "),
        ("Key B", "Value B"),
    ]

    database.replace_specs(
        model_id,
        [
            ("Key B", "Updated"),
            ("Key C", "New"),
        ],
    )

    specs = database.get_specs(model_id)
    assert [(key, value) for _sid, key, value in specs] == [
        ("Key B", "Updated"),
        ("Key C", "New"),
    ]


def test_catalog_entries_include_created_at():
    database.init_db()

    database.add_category("TestCat")
    cats_simple = database.get_categories()
    assert cats_simple and len(cats_simple[0]) == 2
    cats_full = database.get_categories(include_created=True)
    cat_id, _cat_name, cat_created = next(entry for entry in cats_full if entry[1] == "TestCat")
    assert cat_created
    datetime.fromisoformat(cat_created)

    database.add_brand(cat_id, "TestBrand")
    brands_simple = database.get_brands(cat_id)
    assert brands_simple and len(brands_simple[0]) == 2
    brands_full = database.get_brands(cat_id, include_created=True)
    brand_id, _brand_name, brand_created = next(entry for entry in brands_full if entry[1] == "TestBrand")
    assert brand_created
    datetime.fromisoformat(brand_created)

    database.add_model(brand_id, "TestModel")
    models_simple = database.get_models(brand_id)
    assert models_simple and len(models_simple[0]) == 2
    models_full = database.get_models(brand_id, include_created=True)
    model_id, _model_name, model_created = next(entry for entry in models_full if entry[1] == "TestModel")
    assert model_created
    datetime.fromisoformat(model_created)


def test_insert_spec_upserts_duplicate_key():
    model_id = _prepare_model()
    first_id = database.insert_spec(model_id, "Duplicate", "First")
    second_id = database.insert_spec(model_id, "Duplicate", "Second")

    assert first_id == second_id
    specs = database.get_specs(model_id)
    assert specs == [(first_id, "Duplicate", "Second")]


def test_update_spec_duplicate_key_raises_value_error():
    model_id = _prepare_model()
    first_id = database.insert_spec(model_id, "Key1", "Value1")
    second_id = database.insert_spec(model_id, "Key2", "Value2")

    assert first_id and second_id

    with pytest.raises(ValueError):
        database.update_spec(second_id, "Key1", "Updated")

    specs = database.get_specs(model_id)
    assert (first_id, "Key1", "Value1") in specs
    assert (second_id, "Key2", "Value2") in specs


def test_load_specs_map_chunks_large_id_list():
    database.init_db()
    database.add_category("ChunkCat")
    cat_id = next(cid for cid, name in database.get_categories() if name == "ChunkCat")
    database.add_brand(cat_id, "ChunkBrand")
    brand_id = next(bid for bid, name in database.get_brands(cat_id) if name == "ChunkBrand")

    model_ids = []
    for idx in range(5):
        database.add_model(brand_id, f"Model-{idx}")
        mid = next(mid for mid, name in database.get_models(brand_id) if name == f"Model-{idx}")
        database.insert_spec(mid, f"Key-{idx}", f"Value-{idx}")
        model_ids.append(mid)

    many_ids = list(range(1, 2505))
    many_ids.extend(model_ids)

    specs_map = database.load_specs_map(many_ids)

    for idx, mid in enumerate(model_ids):
        assert specs_map[mid][f"Key-{idx}"] == f"Value-{idx}"
