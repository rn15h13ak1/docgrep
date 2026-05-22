import json
from pathlib import Path

import pytest

from config import ConfigError, load_config


def _write_json(p: Path, data) -> str:
    p.write_text(json.dumps(data))
    return str(p)


def test_load_no_file_returns_defaults(tmp_path):
    cfg = load_config(None, default_base_dir=tmp_path)
    assert cfg["search"]["mode"] == "keyword"
    # 既定値の "." が default_base_dir 基準で解決される
    assert cfg["paths"] == [str(tmp_path.resolve())]


def test_backslash_in_paths_raises_config_error(tmp_path):
    p = _write_json(tmp_path / "c.json", {
        "paths": ["C:\\projects\\docs"],
    })
    with pytest.raises(ConfigError) as ei:
        load_config(p)
    assert "paths[0]" in str(ei.value)


def test_backslash_in_output_path_raises(tmp_path):
    p = _write_json(tmp_path / "c.json", {
        "output": {"excel": {"path": "out\\result.xlsx"}},
    })
    with pytest.raises(ConfigError):
        load_config(p)


def test_backslash_in_latest_path_raises(tmp_path):
    p = _write_json(tmp_path / "c.json", {
        "output": {"html": {"latest_path": "out\\latest.html"}},
    })
    with pytest.raises(ConfigError) as ei:
        load_config(p)
    assert "latest_path" in str(ei.value)


def test_html_latest_path_resolved_against_config_dir(tmp_path):
    cfg_path = tmp_path / "c.json"
    _write_json(cfg_path, {
        "output": {"html": {"latest_path": "out/latest.html"}},
    })
    cfg = load_config(str(cfg_path))
    assert cfg["output"]["html"]["latest_path"] == str((tmp_path / "out" / "latest.html").resolve())


def test_relative_paths_resolved_against_config_dir(tmp_path):
    cfg_path = tmp_path / "c.json"
    _write_json(cfg_path, {
        "paths": ["./docs", "../shared"],
        "onenote_export_dir": "./notes",
    })
    cfg = load_config(str(cfg_path))
    assert cfg["paths"][0] == str((tmp_path / "docs").resolve())
    assert cfg["paths"][1] == str((tmp_path.parent / "shared").resolve())
    assert cfg["onenote_export_dir"] == str((tmp_path / "notes").resolve())


def test_absolute_path_passthrough(tmp_path):
    abs_path = "/absolute/elsewhere"
    p = _write_json(tmp_path / "c.json", {"paths": [abs_path]})
    cfg = load_config(p)
    assert cfg["paths"] == [abs_path]


def test_unc_path_passthrough(tmp_path):
    p = _write_json(tmp_path / "c.json", {"paths": ["//server/share/docs"]})
    cfg = load_config(p)
    # Path("//server/share/docs").is_absolute() == True なのでそのまま返る
    assert cfg["paths"][0].startswith("//") or cfg["paths"][0].startswith("/server")


def test_top_level_not_dict_raises(tmp_path):
    p = tmp_path / "c.json"
    p.write_text(json.dumps(["not", "a", "dict"]))
    with pytest.raises(ConfigError):
        load_config(str(p))


def test_invalid_json_raises_config_error(tmp_path):
    p = tmp_path / "c.json"
    p.write_text("{not valid json")
    with pytest.raises(ConfigError):
        load_config(str(p))


def test_deep_merge_keeps_defaults_for_unset_keys(tmp_path):
    p = _write_json(tmp_path / "c.json", {
        "search": {"mode": "regex"},  # 他は既定値を維持
    })
    cfg = load_config(p)
    assert cfg["search"]["mode"] == "regex"
    assert cfg["search"]["operator"] == "and"
    assert cfg["search"]["case_sensitive"] is False
