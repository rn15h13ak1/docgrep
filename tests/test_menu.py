"""menu.py のヘルパ関数テスト（subprocess.run はモック）。"""
import json
import subprocess as _subprocess
import sys
import types
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

import menu


# ------------------------------------------------------------ history I/O ----

def test_load_history_empty_when_missing(tmp_path, monkeypatch):
    monkeypatch.setattr(menu, "HISTORY_PATH", tmp_path / "no.json")
    assert menu._load_history() == {}


def test_save_then_load_history(tmp_path, monkeypatch):
    p = tmp_path / "hist.json"
    monkeypatch.setattr(menu, "HISTORY_PATH", p)
    menu._save_history({"keywords": "東京 プロジェクト",
                        "mode": "fuzzy", "paths": ["/x", "/y"]})
    assert p.is_file()
    data = json.loads(p.read_text(encoding="utf-8"))
    assert data["keywords"] == "東京 プロジェクト"
    assert data["mode"] == "fuzzy"
    assert data["paths"] == ["/x", "/y"]
    # _load_history でも同じ内容が読める
    loaded = menu._load_history()
    assert loaded == data


def test_load_history_corrupted_file_returns_empty(tmp_path, monkeypatch):
    p = tmp_path / "bad.json"
    p.write_text("{not json", encoding="utf-8")
    monkeypatch.setattr(menu, "HISTORY_PATH", p)
    assert menu._load_history() == {}


def test_save_history_failure_is_silent(monkeypatch, tmp_path):
    # 書き込み不可なパスでも例外が漏れない
    bad = tmp_path / "missing_dir" / "h.json"
    monkeypatch.setattr(menu, "HISTORY_PATH", bad)
    # 通るだけで OK
    menu._save_history({"keywords": "x"})


# ----------------------------------------------------- _ask_with_history ----

def test_ask_with_history_uses_prev_on_empty(monkeypatch):
    monkeypatch.setattr("builtins.input", lambda *_: "")
    assert menu._ask_with_history("prompt", "前回値") == "前回値"


def test_ask_with_history_returns_new_input(monkeypatch):
    monkeypatch.setattr("builtins.input", lambda *_: "新規")
    assert menu._ask_with_history("prompt", "前回値") == "新規"


def test_ask_with_history_no_prev_returns_empty(monkeypatch):
    monkeypatch.setattr("builtins.input", lambda *_: "")
    assert menu._ask_with_history("prompt", None) == ""


# ---------------------------------------------------- _resolve_latest_path ---

def test_resolve_latest_path_from_explicit_config(tmp_path):
    yaml_text = (
        "output:\n"
        "  html:\n"
        '    latest_path: "./reports/latest.html"\n'
    )
    cfg = tmp_path / "config.yaml"
    cfg.write_text(yaml_text, encoding="utf-8")
    resolved = menu._resolve_latest_html_path(cfg)
    assert resolved == (tmp_path / "reports" / "latest.html").resolve()


def test_resolve_latest_path_returns_none_when_no_setting(tmp_path):
    cfg = tmp_path / "config.yaml"
    cfg.write_text("paths:\n  - .\n", encoding="utf-8")
    assert menu._resolve_latest_html_path(cfg) is None


def test_resolve_latest_path_missing_config(tmp_path):
    assert menu._resolve_latest_html_path(tmp_path / "nope.yaml") is None


# ------------------------------------------- _run_docgrep / _run_export -----

def test_run_docgrep_invokes_subprocess_with_expected_args(monkeypatch):
    captured = {}

    def fake_run(cmd, cwd=None, **kw):
        captured["cmd"] = cmd
        captured["cwd"] = cwd
        return MagicMock(returncode=0)

    monkeypatch.setattr(menu.subprocess, "run", fake_run)
    monkeypatch.setattr("builtins.input", lambda *_: "")

    rc = menu._run_docgrep(["foo", "--mode", "keyword"], wait=False)
    assert rc == 0
    cmd = captured["cmd"]
    # 先頭が python 実行ファイル、続いて docgrep.py と引数
    assert cmd[1] == "docgrep.py"
    assert "foo" in cmd
    assert "--mode" in cmd and "keyword" in cmd
    assert captured["cwd"] == menu.SCRIPT_DIR


def test_run_export_invokes_powershell(monkeypatch):
    calls = []

    def fake_run(cmd, cwd=None, **kw):
        calls.append(cmd)
        return MagicMock(returncode=0)

    monkeypatch.setattr(menu.subprocess, "run", fake_run)
    monkeypatch.setattr("builtins.input", lambda *_: "")

    rc = menu._run_export(wait=False)
    assert rc == 0

    # 2 回呼ばれる: 1) Unblock-File 試行, 2) 本体起動
    assert len(calls) >= 1
    main_cmd = calls[-1]
    assert main_cmd[0] == "powershell"
    assert "-NoProfile" in main_cmd
    assert "-ExecutionPolicy" in main_cmd and "Bypass" in main_cmd
    assert "-File" in main_cmd
    # スクリプト本体のパスを最後の要素として渡している
    assert main_cmd[-1].endswith("export_onenote.ps1")


def test_run_docgrep_reports_each_exit_code(monkeypatch, capsys):
    """0/1/130/その他 で出力メッセージが切り替わること。"""
    for rc_in, expected_substring in [
        (0, "ヒットあり"),
        (1, "ヒットなし"),
        (130, "中断"),
        (2, "エラー"),
    ]:
        monkeypatch.setattr(
            menu.subprocess, "run",
            lambda cmd, cwd=None, **kw: MagicMock(returncode=rc_in),
        )
        menu._run_docgrep([], wait=False)
        out = capsys.readouterr().out
        assert expected_substring in out, \
            f"rc={rc_in} で {expected_substring!r} が出力に含まれない: {out}"
