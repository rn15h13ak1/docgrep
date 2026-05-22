#!/usr/bin/env python3
"""
docgrep 対話メニュー
====================
docgrep の全文検索・OneNote エクスポートを対話形式で実行する。

使い方:
  python menu.py
"""

import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

SCRIPT_DIR = Path(__file__).parent.resolve()
WIDTH = 62

# 入力履歴ファイル（前回キーワード / モード / パスを覚えて再利用）
HISTORY_PATH = Path.home() / ".docgrep_history.json"


# ================================================================
# UI ユーティリティ
# ================================================================

def hr(char="="):
    print(char * WIDTH)


def print_menu(title, items, back_label="戻る"):
    """メニューを表示して選択番号を返す。0 = 戻る / 終了"""
    while True:
        print()
        hr()
        print(f"  {title}")
        hr()
        for i, item in enumerate(items, 1):
            print(f"  {i}. {item}")
        hr("-")
        print(f"  0. {back_label}")
        hr()
        choice = input("番号を入力してください: ").strip()
        if choice == "0":
            return 0
        if choice.isdigit() and 1 <= int(choice) <= len(items):
            return int(choice)
        print("  ※ 無効な入力です。もう一度入力してください。")


def ask_yes_no(prompt: str, default: bool = False) -> bool:
    """y/n 質問。空 Enter で default。"""
    suffix = " [Y/n]" if default else " [y/N]"
    while True:
        ans = input(f"  {prompt}{suffix}: ").strip().lower()
        if not ans:
            return default
        if ans in ("y", "yes"):
            return True
        if ans in ("n", "no"):
            return False
        print("  ※ y か n を入力してください。")


def ask_float(prompt: str, default: Optional[float] = None,
              min_val: float = 0.0, max_val: float = 1.0) -> Optional[float]:
    """範囲付き float 入力。Enter で default。"""
    suffix = f" [{default}]" if default is not None else ""
    while True:
        ans = input(f"  {prompt}{suffix}: ").strip()
        if not ans:
            return default
        try:
            v = float(ans)
        except ValueError:
            print(f"  ※ 数値を入力してください ({min_val}〜{max_val})")
            continue
        if not (min_val <= v <= max_val):
            print(f"  ※ {min_val}〜{max_val} の範囲で指定してください")
            continue
        return v


def wait_enter():
    input("\n  Enter キーでメニューに戻ります...")


# ================================================================
# 履歴の読み書き
# ================================================================

def _load_history() -> Dict[str, Any]:
    """前回の入力履歴を読み込む。読み込み失敗時は空 dict。"""
    if not HISTORY_PATH.is_file():
        return {}
    try:
        with HISTORY_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_history(entry: Dict[str, Any]) -> None:
    """履歴を保存する。失敗しても致命ではないので握りつぶす。"""
    try:
        HISTORY_PATH.write_text(
            json.dumps(entry, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def _ask_with_history(prompt: str, prev: Optional[str]) -> str:
    """前回値があれば [前回値] を表示し、空 Enter で再利用させる。"""
    label = f"{prompt} [{prev}]" if prev else prompt
    ans = input(f"  {label}: ").strip()
    if not ans and prev:
        return prev
    return ans


# ================================================================
# config.yaml から output.html.latest_path を取得
# ================================================================

def _resolve_latest_html_path(config_path: Optional[Path]) -> Optional[Path]:
    """config.yaml の output.html.latest_path を解決して返す。

    config 未指定なら CWD → SCRIPT_DIR の順で探す。読めなければ None。
    """
    try:
        import yaml  # type: ignore
    except ImportError:
        return None

    candidates = []
    if config_path:
        candidates.append(Path(config_path))
    else:
        candidates.append(Path.cwd() / "config.yaml")
        candidates.append(SCRIPT_DIR / "config.yaml")

    cfg_file = next((p for p in candidates if p.is_file()), None)
    if not cfg_file:
        return None

    try:
        with cfg_file.open("r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
    except Exception:
        return None

    latest = cfg.get("output", {}).get("html", {}).get("latest_path")
    if not isinstance(latest, str) or not latest:
        return None

    p = Path(latest)
    if not p.is_absolute():
        p = (cfg_file.resolve().parent / p).resolve()
    return p


def _open_in_browser(path: Path) -> None:
    try:
        if sys.platform == "win32":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
        print(f"  HTML レポートを開きました: {path}")
    except Exception as e:
        print(f"  ※ レポートを開けませんでした: {e}")
        print(f"     直接開いてください: {path}")


# ================================================================
# 子プロセス起動
# ================================================================

def _run_docgrep(args: List[str], wait: bool = True) -> int:
    """SCRIPT_DIR/docgrep.py を subprocess で実行する。"""
    cmd = [sys.executable, "docgrep.py"] + args

    print()
    cmd_str = " ".join(["python", "docgrep.py"] + args)
    print(f"  実行: {cmd_str}")
    hr("-")

    result = subprocess.run(cmd, cwd=SCRIPT_DIR)

    hr("-")
    if result.returncode == 0:
        print("  完了しました（ヒットあり）。")
    elif result.returncode == 1:
        print("  完了しました（ヒットなし）。")
    elif result.returncode == 130:
        print("  中断されました。")
    else:
        print(f"  エラーが発生しました（終了コード: {result.returncode}）")

    if wait:
        wait_enter()
    return result.returncode


def _run_export(wait: bool = True) -> int:
    """OneNote を Word(.docx) に一括エクスポート（粒度は ps1 既定の section 固定）。"""
    script = SCRIPT_DIR / "export_onenote.ps1"
    cmd = ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(script)]

    print()
    print(f"  実行: powershell -ExecutionPolicy Bypass -File export_onenote.ps1")
    hr("-")

    result = subprocess.run(cmd, cwd=SCRIPT_DIR)

    hr("-")
    if result.returncode == 0:
        print("  完了しました。")
    else:
        print(f"  エラーが発生しました（終了コード: {result.returncode}）")

    if wait:
        wait_enter()
    return result.returncode


# ================================================================
# 全文検索フロー
# ================================================================

def _input_paths() -> List[str]:
    print()
    print("  検索パスを入力してください（1行1パス、空Enter で確定）。")
    paths = []
    while True:
        line = input(f"  パス{len(paths) + 1}: ").strip()
        if not line:
            break
        paths.append(line)
    return paths


def _input_config() -> Tuple[Optional[str], Optional[Path]]:
    """設定ファイル選択。返り値: (cli 引数用 path or None, 解決済み Path or None)"""
    choice = print_menu(
        "docgrep - 設定ファイル",
        ["既定（CWD → スクリプト同梱の config.yaml）", "別ファイルを指定"],
    )
    if choice == 0:
        return None, None
    if choice == 1:
        # 既定 → CLI 引数なし、自動検出に任せる
        return "__default__", None
    ans = input("  設定ファイルのパス: ").strip()
    if not ans:
        return None, None
    p = Path(ans).expanduser()
    if not p.is_file():
        print(f"  ※ ファイルが見つかりません: {p}")
        wait_enter()
        return None, None
    return str(p), p.resolve()


def _ask_mode_options(mode: str, n_keywords: int) -> Dict[str, str]:
    """モード固有のオプションを尋ねる。返り値は CLI 引数の dict（key=value）形式。"""
    opts: Dict[str, str] = {}
    if mode == "keyword" and n_keywords >= 2:
        choice = print_menu(
            "複数キーワードの結合",
            ["AND（すべて含む）", "OR（いずれか含む）"],
            back_label="既定 (AND)",
        )
        if choice == 1:
            opts["--operator"] = "and"
        elif choice == 2:
            opts["--operator"] = "or"
    elif mode == "fuzzy":
        threshold = ask_float(
            "あいまい検索のしきい値 (0.0-1.0、Enter で既定 0.80)",
            default=None,
        )
        if threshold is not None:
            opts["--fuzzy-threshold"] = f"{threshold:g}"
    return opts


def _ask_common_options() -> List[str]:
    """共通オプション（大小区別 / NFKC / verbose）を任意指定。"""
    if not ask_yes_no("詳細オプションを指定しますか？", default=False):
        return []
    flags: List[str] = []
    if ask_yes_no("大文字小文字を区別する？", default=False):
        flags.append("--case-sensitive")
    if not ask_yes_no("全角半角の正規化を有効にする？（既定: 有効）", default=True):
        flags.append("--no-normalize-width")
    if ask_yes_no("詳細ログ (--verbose) を出力する？", default=False):
        flags.append("--verbose")
    return flags


def _run_search():
    history = _load_history()
    prev_keywords = history.get("keywords", "")
    prev_mode = history.get("mode", "")
    prev_paths: List[str] = history.get("paths", []) or []

    while True:
        # Step 0: 設定ファイル選択
        config_arg, resolved_config = _input_config()
        if config_arg is None:
            return  # 中止

        # Step 1: 検索パス選択（前回履歴があれば「前回のパスを再利用」も提示）
        path_options = ["config.yaml の設定に従う", "パスを指定する"]
        if prev_paths:
            path_options.append(f"前回のパスを再利用 ({len(prev_paths)} 件)")
        path_choice = print_menu("docgrep - 検索パス", path_options)
        if path_choice == 0:
            return

        custom_paths: List[str] = []
        if path_choice == 2:
            custom_paths = _input_paths()
            if not custom_paths:
                print("  パスが入力されていません。中止します。")
                wait_enter()
                return
        elif path_choice == 3 and prev_paths:
            custom_paths = list(prev_paths)
            print(f"  前回のパスを再利用: {custom_paths}")

        # Step 2: 検索キーワード入力（前回値を [履歴] で提示）
        print()
        keywords_str = _ask_with_history(
            "検索キーワード（スペース区切り、空で中止）", prev_keywords
        )
        if not keywords_str:
            return
        keywords = keywords_str.split()

        # Step 3: 検索モード選択（前回値があれば既定として 0=戻る ではなく直接採用）
        mode_items = [
            "キーワード検索（keyword）",
            "正規表現検索（regex）",
            "あいまい検索（fuzzy）",
        ]
        prev_mode_idx = {"keyword": 1, "regex": 2, "fuzzy": 3}.get(prev_mode)
        if prev_mode_idx:
            mode_items[prev_mode_idx - 1] += "  ← 前回"
        mode_choice = print_menu("docgrep - 検索モード", mode_items)
        if mode_choice == 0:
            return
        mode_map = {1: "keyword", 2: "regex", 3: "fuzzy"}
        mode = mode_map[mode_choice]

        # Step 4: モード固有オプション
        mode_opts = _ask_mode_options(mode, len(keywords))

        # Step 5: 共通オプション
        common_flags = _ask_common_options()

        # Step 6: 入力内容確認
        print()
        hr()
        print("  入力内容の確認")
        hr()
        print(f"  設定ファイル   : {resolved_config or '既定（自動検出）'}")
        if custom_paths:
            print("  検索パス       :")
            for i, p in enumerate(custom_paths, 1):
                print(f"    [{i}] {p}")
        else:
            print("  検索パス       : config.yaml の設定に従う")
        print(f"  検索キーワード : {' '.join(keywords)}")
        print(f"  検索モード     : {mode}")
        for k, v in mode_opts.items():
            print(f"  {k:14}: {v}")
        if common_flags:
            print(f"  追加フラグ     : {' '.join(common_flags)}")
        hr()

        confirm_choice = print_menu(
            "実行確認",
            ["この内容で実行", "最初から入力し直す"],
            back_label="中止",
        )
        if confirm_choice == 0:
            return
        if confirm_choice == 2:
            continue

        # 実行直前に履歴を保存（次回起動で再利用）
        _save_history({
            "keywords": " ".join(keywords),
            "mode": mode,
            "paths": custom_paths,
        })

        # Step 7: 実行
        args: List[str] = list(keywords)
        if config_arg != "__default__":
            args.extend(["-c", config_arg])
        args.extend(["--mode", mode])
        for k, v in mode_opts.items():
            args.extend([k, v])
        args.extend(common_flags)
        for p in custom_paths:
            args.extend(["-p", p])

        rc = _run_docgrep(args, wait=False)

        # Step 8: レポートを開くか（ヒットあり/なしでも参照したい場合がある）
        if rc in (0, 1):
            latest = _resolve_latest_html_path(resolved_config)
            if latest and latest.is_file():
                if ask_yes_no("HTML レポートを開きますか？", default=True):
                    _open_in_browser(latest)
        wait_enter()
        return


# ================================================================
# メインメニュー
# ================================================================

def main() -> int:
    options = [
        "全文検索を実行",
        "OneNote エクスポート（docgrep 用前処理）",
        "OneNote エクスポート → 全文検索（連続実行）",
    ]
    while True:
        choice = print_menu("docgrep", options, back_label="終了")
        if choice == 0:
            print("\n  終了します。\n")
            return 0
        if choice == 1:
            _run_search()
        elif choice == 2:
            _run_export(wait=True)
        elif choice == 3:
            rc = _run_export(wait=False)
            if rc == 0:
                _run_search()
            else:
                print(f"\n  エクスポートが失敗したため検索を中止します（終了コード: {rc}）")
                wait_enter()


if __name__ == "__main__":
    sys.exit(main() or 0)
