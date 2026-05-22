#!/usr/bin/env python3
"""
docgrep 対話メニュー
====================
docgrep の全文検索・OneNote エクスポートを対話形式で実行する。

使い方:
  python menu.py
"""

import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
WIDTH = 62


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


def wait_enter():
    input("\n  Enter キーでメニューに戻ります...")


# ================================================================
# 子プロセス起動
# ================================================================

def _run_docgrep(args: list, wait: bool = True) -> int:
    """SCRIPT_DIR/docgrep.py を subprocess で実行する。"""
    cmd = [sys.executable, "docgrep.py"] + args

    print()
    cmd_str = " ".join(["python", "docgrep.py"] + args)
    print(f"  実行: {cmd_str}")
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


def _run_export(wait: bool = True) -> int:
    """
    OneNote を Word(.docx) に一括エクスポート（粒度は ps1 既定の section 固定）。
    戻り値: PowerShell スクリプトの終了コード
    """
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

def _input_paths() -> list:
    """パスを複数入力させる（空Enter で終了）"""
    print()
    print("  検索パスを入力してください（1行1パス、空Enter で確定）。")
    paths = []
    while True:
        line = input(f"  パス{len(paths) + 1}: ").strip()
        if not line:
            break
        paths.append(line)
    return paths


def _run_search():
    """
    全文検索を実行する。
    フロー:
      1. 検索パスを config.yaml に従うか手動指定か選択
      2. 検索キーワードを入力
      3. 検索モードを選択
      4. 入力内容を確認（実行 / 最初から入力し直し / 中止）
      5. 実行
    """
    while True:
        # Step 1: 検索パス選択
        path_choice = print_menu(
            "docgrep - 検索パス",
            ["config.yaml の設定に従う", "パスを指定する"],
        )
        if path_choice == 0:
            return

        custom_paths = []
        if path_choice == 2:
            custom_paths = _input_paths()
            if not custom_paths:
                print("  パスが入力されていません。中止します。")
                wait_enter()
                return

        # Step 2: 検索キーワード入力
        print()
        keywords_str = input("  検索キーワード（スペース区切り、空で中止）: ").strip()
        if not keywords_str:
            return
        keywords = keywords_str.split()

        # Step 3: 検索モード選択
        mode_items = [
            "キーワード検索（keyword）",
            "正規表現検索（regex）",
            "あいまい検索（fuzzy）",
        ]
        mode_choice = print_menu("docgrep - 検索モード", mode_items)
        if mode_choice == 0:
            return
        mode_map = {1: "keyword", 2: "regex", 3: "fuzzy"}
        mode = mode_map[mode_choice]

        # Step 4: 入力内容の確認
        print()
        hr()
        print("  入力内容の確認")
        hr()
        if custom_paths:
            print("  検索パス       :")
            for i, p in enumerate(custom_paths, 1):
                print(f"    [{i}] {p}")
        else:
            print("  検索パス       : config.yaml の設定に従う")
        print(f"  検索キーワード : {' '.join(keywords)}")
        print(f"  検索モード     : {mode}")
        hr()

        confirm_choice = print_menu(
            "実行確認",
            ["この内容で実行", "最初から入力し直す"],
            back_label="中止",
        )
        if confirm_choice == 0:
            return
        if confirm_choice == 2:
            continue  # Step 1 に戻る

        # Step 5: 実行
        args = keywords + ["--mode", mode]
        for p in custom_paths:
            args.extend(["-p", p])
        _run_docgrep(args)
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
