"""起動時セルフチェック。

- 必須同梱パッケージの import 可否
- MS Office (COM) の起動可否（Windows のみ）
- PDF / OneNote(.one) はスコープ外として注意表示
"""
from __future__ import annotations

import sys
from dataclasses import dataclass, field
from typing import List


REQUIRED_PACKAGES = [
    ("openpyxl", "openpyxl"),
    ("lxml", "lxml"),
    ("charset_normalizer", "charset-normalizer"),
    ("psutil", "psutil"),
    ("tqdm", "tqdm"),
    ("jinja2", "jinja2"),
    ("yaml", "PyYAML"),
]


@dataclass
class CheckResult:
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    infos: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors


def _check_packages(result: CheckResult) -> None:
    missing = []
    for mod, dist in REQUIRED_PACKAGES:
        try:
            __import__(mod)
        except ImportError:
            missing.append(dist)
    if missing:
        result.errors.append(
            "必須パッケージが不足: " + ", ".join(missing)
            + "。Anaconda（フル版）環境での実行を確認してください。"
        )
    else:
        result.infos.append(
            "必須パッケージ OK（openpyxl/lxml/charset-normalizer/jinja2/pywin32 等）"
        )


def _check_office(result: CheckResult, require_office: bool) -> None:
    if sys.platform != "win32":
        msg = (
            "Windows 以外のOSのため MS Office COM は利用できません。"
            "Office文書(.doc/.docx/.ppt/.pptx/.xls)は抽出対象から除外されます。"
        )
        if require_office:
            result.errors.append(
                msg + " 本ツールは Microsoft Office インストール済み Windows での実行を前提とします。"
                " テキスト/.xlsx のみ検索する場合は runtime.require_office=false を指定してください。"
            )
        else:
            result.warnings.append(msg)
        return

    try:
        import win32com.client  # type: ignore  # noqa: F401
    except ImportError:
        msg = "pywin32（win32com）がインポートできません。Office 文書の抽出に必要です。"
        if require_office:
            result.errors.append(msg)
        else:
            result.warnings.append(msg + "（Office 文書はスキップします）")
        return

    issues = []
    for app_name in ("Word.Application", "Excel.Application", "PowerPoint.Application"):
        try:
            import win32com.client as _w  # type: ignore
            obj = _w.DispatchEx(app_name)
            try:
                obj.Visible = False
            except Exception:
                pass
            try:
                obj.Quit()
            except Exception:
                pass
        except Exception as e:  # COM 起動失敗
            issues.append(f"{app_name}: {e}")
    if issues:
        msg = "MS Office (COM) の起動に失敗: " + " / ".join(issues)
        if require_office:
            result.errors.append(msg)
        else:
            result.warnings.append(msg + "（Office 文書はスキップします）")
    else:
        result.infos.append("MS Office (COM) OK（前提条件）")


def run_selfcheck(require_office: bool) -> CheckResult:
    result = CheckResult()
    _check_packages(result)
    _check_office(result, require_office)
    # スコープ外形式（参考表示）
    result.warnings.append(
        "PDF / OneNote(.one) は初版スコープ外のため該当ファイルはスキップします。"
        "（OneNote は同梱の export_onenote.ps1 で Word 化すれば通常検索されます）"
    )
    return result


def print_result(result: CheckResult) -> None:
    print("[セルフチェック]")
    for info in result.infos:
        print(f"  [OK]   {info}")
    for w in result.warnings:
        print(f"  [WARN] {w}")
    for e in result.errors:
        print(f"  [NG]   {e}")
    if result.errors:
        print("  → 前提条件違反のため起動を中止します。")
    else:
        print("  → 検索を実行します。")
