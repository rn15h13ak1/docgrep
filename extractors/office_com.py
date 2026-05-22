"""MS Office COM (pywin32) 経由の抽出: Word / PowerPoint / 旧 Excel(.xls)。

- 不可視・専用インスタンスを新規生成し、ユーザーが開いている既存ドキュメントには触れない
- 対象ファイルは ReadOnly=True で開き、書き込みロックを取らない
- 一定件数ごとにインスタンスを再生成してメモリリークを抑える
- 確実に Close / Quit する
- 各抽出器は List[Segment] を返し、locator にシート名・スライド番号・図形名等を入れる

内部設計:
  _AppHandle: Word/Excel/PowerPoint で共通の DispatchEx / Quit / recycle カウンタを集約
  _WordHandle / _ExcelHandle / _PptHandle: アプリ固有の起動オプションだけを上書き
  OfficeCom: 3 つのハンドルをまとめて持ち、extract_* メソッドを公開する
"""
from __future__ import annotations

import gc
import os
import sys
from typing import Dict, List

from search import Segment

try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    from pywintypes import com_error  # type: ignore
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    com_error = Exception  # type: ignore


def is_available() -> bool:
    """COM 抽出が利用可能か（Windows かつ pywin32 がある）。"""
    return sys.platform == "win32" and HAS_WIN32


# =============================================================================
# COM ハンドル共通基底
# =============================================================================

class _AppHandle:
    """Office アプリの COM ハンドル管理（DispatchEx + Quit + 再生成カウンタ）。

    サブクラスで `PROG_ID` と `_configure(app)` を上書きする。
    """
    PROG_ID: str = ""

    def __init__(self, recycle_every: int) -> None:
        self.app = None
        self.count = 0
        self.recycle_every = max(1, int(recycle_every))

    def get(self):
        if self.app is None:
            self.app = win32com.client.DispatchEx(self.PROG_ID)
            self._configure(self.app)
        return self.app

    def _configure(self, app) -> None:
        """アプリ固有の起動オプション (Visible=False など) を設定する。"""

    def tick(self) -> None:
        """1 ファイル処理ごとに呼ぶ。recycle_every 件で破棄して再生成可能に。"""
        self.count += 1
        if self.count >= self.recycle_every:
            self.recycle()

    def recycle(self) -> None:
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
            self.app = None
        self.count = 0
        gc.collect()


class _WordHandle(_AppHandle):
    PROG_ID = "Word.Application"

    def _configure(self, app) -> None:
        try:
            app.Visible = False
        except Exception:
            pass
        try:
            app.DisplayAlerts = 0  # wdAlertsNone
        except Exception:
            pass


class _ExcelHandle(_AppHandle):
    PROG_ID = "Excel.Application"

    def _configure(self, app) -> None:
        try:
            app.Visible = False
        except Exception:
            pass
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        try:
            app.AskToUpdateLinks = False
        except Exception:
            pass


class _PptHandle(_AppHandle):
    PROG_ID = "PowerPoint.Application"

    def _configure(self, app) -> None:
        # PowerPoint は Visible=False を受け付けない場合がある
        try:
            app.DisplayAlerts = 1  # ppAlertsNone
        except Exception:
            pass


# =============================================================================
# OfficeCom: Word/Excel/PowerPoint をまとめて提供する抽出器
# =============================================================================

class OfficeCom:
    """Word/Excel(.xls)/PowerPoint の COM インスタンスを管理する抽出器。"""

    def __init__(self, recycle_every: int = 30) -> None:
        if not is_available():
            raise RuntimeError("OfficeCom is not available on this platform.")
        self._word = _WordHandle(recycle_every)
        self._excel = _ExcelHandle(recycle_every)
        self._ppt = _PptHandle(recycle_every)
        pythoncom.CoInitialize()

    # 旧 API 互換: 外部から self.counts を読まれてもよいように property で公開
    @property
    def counts(self) -> Dict[str, int]:
        return {
            "word": self._word.count,
            "excel": self._excel.count,
            "ppt": self._ppt.count,
        }

    def __enter__(self) -> "OfficeCom":
        return self

    def __exit__(self, *exc) -> None:
        self.shutdown()

    def shutdown(self) -> None:
        for h in (self._word, self._excel, self._ppt):
            h.recycle()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def recover_all(self) -> None:
        """COM が応答しなくなったときに全インスタンスを破棄して再生成可能な状態にする。

        タイムアウト超過などで Office プロセスが応答しなくなった疑いがあるときに
        呼ぶ。次回 get() で新規 DispatchEx される。Quit() 自体が固まる可能性が
        あるため例外は全て無視する。
        """
        for h in (self._word, self._excel, self._ppt):
            h.recycle()

    # =========================================================
    # Word
    # =========================================================
    def extract_word(self, path: str) -> List[Segment]:
        word = self._word.get()
        doc = None
        segments: List[Segment] = []
        try:
            doc = word.Documents.Open(
                FileName=os.path.abspath(path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument="",
                Visible=False,
            )
            # 本文
            try:
                body = doc.Content.Text or ""
                if body.strip():
                    segments.append(Segment(text=body, locator="本文"))
            except com_error:
                pass
            # ヘッダ/フッタ
            try:
                for i, section in enumerate(doc.Sections, start=1):
                    try:
                        for hf in section.Headers:
                            if hf.Exists:
                                txt = hf.Range.Text or ""
                                if txt.strip():
                                    segments.append(Segment(
                                        text=txt,
                                        locator=f"セクション {i} / ヘッダー",
                                    ))
                        for hf in section.Footers:
                            if hf.Exists:
                                txt = hf.Range.Text or ""
                                if txt.strip():
                                    segments.append(Segment(
                                        text=txt,
                                        locator=f"セクション {i} / フッター",
                                    ))
                    except com_error:
                        continue
            except com_error:
                pass
            # 図形内テキスト
            try:
                for shape in doc.Shapes:
                    try:
                        tf = shape.TextFrame
                        if tf.HasText:
                            txt = tf.TextRange.Text or ""
                            if txt.strip():
                                name = ""
                                try:
                                    name = shape.Name or ""
                                except com_error:
                                    pass
                                loc = f"図形「{name}」" if name else "図形"
                                segments.append(Segment(text=txt, locator=loc))
                    except com_error:
                        continue
            except com_error:
                pass
            # コメント
            try:
                for i, c in enumerate(doc.Comments, start=1):
                    try:
                        txt = c.Range.Text or ""
                        if not txt.strip():
                            continue
                        author = ""
                        try:
                            author = c.Author or ""
                        except com_error:
                            pass
                        loc = f"コメント #{i}"
                        if author:
                            loc += f" ({author})"
                        segments.append(Segment(text=txt, locator=loc))
                    except com_error:
                        continue
            except com_error:
                pass
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
            self._word.tick()
        return segments

    # =========================================================
    # PowerPoint
    # =========================================================
    def extract_powerpoint(self, path: str) -> List[Segment]:
        app = self._ppt.get()
        pres = None
        segments: List[Segment] = []
        try:
            pres = app.Presentations.Open(
                FileName=os.path.abspath(path),
                ReadOnly=True,
                Untitled=False,
                WithWindow=False,
            )
            for slide in pres.Slides:
                try:
                    slide_no = int(slide.SlideNumber)
                except com_error:
                    slide_no = 0
                # スライド内の図形
                try:
                    for shape in slide.Shapes:
                        try:
                            if shape.HasTextFrame and shape.TextFrame.HasText:
                                txt = shape.TextFrame.TextRange.Text or ""
                                if not txt.strip():
                                    continue
                                shape_name = ""
                                try:
                                    shape_name = shape.Name or ""
                                except com_error:
                                    pass
                                loc = f"スライド {slide_no} / {shape_name}" if shape_name \
                                      else f"スライド {slide_no}"
                                segments.append(Segment(text=txt, locator=loc))
                        except com_error:
                            continue
                except com_error:
                    pass
                # ノート
                try:
                    if slide.HasNotesPage:
                        for shape in slide.NotesPage.Shapes:
                            try:
                                if shape.HasTextFrame and shape.TextFrame.HasText:
                                    txt = shape.TextFrame.TextRange.Text or ""
                                    if txt.strip():
                                        segments.append(Segment(
                                            text=txt,
                                            locator=f"スライド {slide_no} / ノート",
                                        ))
                            except com_error:
                                continue
                except com_error:
                    pass
        finally:
            if pres is not None:
                try:
                    pres.Close()
                except Exception:
                    pass
            self._ppt.tick()
        return segments

    # =========================================================
    # Excel (旧形式 .xls)
    # =========================================================
    def extract_excel_old(self, path: str) -> List[Segment]:
        excel = self._excel.get()
        wb = None
        segments: List[Segment] = []
        try:
            wb = excel.Workbooks.Open(
                Filename=os.path.abspath(path),
                ReadOnly=True,
                UpdateLinks=0,
                AddToMru=False,
                IgnoreReadOnlyRecommended=True,
            )
            for ws in wb.Worksheets:
                try:
                    sheet_name = str(ws.Name)
                except Exception:
                    sheet_name = ""
                if sheet_name:
                    segments.append(Segment(
                        text=sheet_name,
                        locator=f"シート名: {sheet_name}",
                    ))
                # セル値（UsedRange）— Value は 2 次元タプル、座標は UsedRange.Row/Column 基点
                try:
                    used = ws.UsedRange
                    val = used.Value
                    start_row = int(used.Row)
                    start_col = int(used.Column)
                    if val is None:
                        pass
                    elif isinstance(val, tuple) and val and isinstance(val[0], tuple):
                        # 2 次元
                        for i, row in enumerate(val):
                            for j, c in enumerate(row):
                                if c is None:
                                    continue
                                coord = _col_letter(start_col + j) + str(start_row + i)
                                segments.append(Segment(
                                    text=str(c),
                                    locator=f"{sheet_name}!{coord}",
                                ))
                    elif isinstance(val, tuple):
                        # 1 行/列のみ（タプル）
                        for j, c in enumerate(val):
                            if c is None:
                                continue
                            coord = _col_letter(start_col + j) + str(start_row)
                            segments.append(Segment(
                                text=str(c),
                                locator=f"{sheet_name}!{coord}",
                            ))
                    else:
                        coord = _col_letter(start_col) + str(start_row)
                        segments.append(Segment(
                            text=str(val),
                            locator=f"{sheet_name}!{coord}",
                        ))
                except com_error:
                    pass
                # 図形内テキスト
                try:
                    for shape in ws.Shapes:
                        try:
                            tf = shape.TextFrame
                            text = tf.Characters().Text
                            if not text or not text.strip():
                                continue
                            shape_name = ""
                            try:
                                shape_name = shape.Name or ""
                            except com_error:
                                pass
                            prefix = f"{sheet_name}: " if sheet_name else ""
                            loc = f"{prefix}図形「{shape_name}」" if shape_name \
                                  else f"{prefix}図形"
                            segments.append(Segment(text=text, locator=loc))
                        except (com_error, AttributeError):
                            continue
                except com_error:
                    pass
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            self._excel.tick()
        return segments


def _col_letter(col_num: int) -> str:
    """1 始まりの列番号を Excel の列記号 (A, B, ..., AA, AB, ...) に変換する。"""
    if col_num < 1:
        return ""
    s = ""
    n = col_num
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s
