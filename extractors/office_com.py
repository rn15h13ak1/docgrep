"""MS Office COM (pywin32) 経由の抽出: Word / PowerPoint / 旧 Excel(.xls)。

- 不可視・専用インスタンスを新規生成し、ユーザーが開いている既存ドキュメントには触れない
- 対象ファイルは ReadOnly=True で開き、書き込みロックを取らない
- 一定件数ごとにインスタンスを再生成してメモリリークを抑える
- 確実に Close / Quit する
- 各抽出器は List[Segment] を返し、locator にシート名・スライド番号・図形名等を入れる
"""
from __future__ import annotations

import gc
import os
import sys
from typing import List

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


class OfficeCom:
    """Word/Excel(.xls)/PowerPoint の COM インスタンスを管理する抽出器。"""

    def __init__(self, recycle_every: int = 30) -> None:
        if not is_available():
            raise RuntimeError("OfficeCom is not available on this platform.")
        self.recycle_every = max(1, int(recycle_every))
        self.word = None
        self.excel = None
        self.ppt = None
        self.counts = {"word": 0, "excel": 0, "ppt": 0}
        pythoncom.CoInitialize()

    def __enter__(self) -> "OfficeCom":
        return self

    def __exit__(self, *exc) -> None:
        self.shutdown()

    def shutdown(self) -> None:
        for attr in ("word", "excel", "ppt"):
            app = getattr(self, attr, None)
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
                setattr(self, attr, None)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def _recycle(self, key: str) -> None:
        app = getattr(self, key, None)
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
            setattr(self, key, None)
        self.counts[key] = 0
        gc.collect()

    def recover_all(self) -> None:
        """COM が応答しなくなったときに全インスタンスを破棄して再生成可能な状態にする。

        タイムアウト超過などで Office プロセスが応答しなくなった疑いがあるときに
        呼ぶ。次回 _get_* で新規 DispatchEx される。Quit() 自体が固まる可能性が
        あるため、wait 系の例外は全て無視する。
        """
        for key in ("word", "excel", "ppt"):
            app = getattr(self, key, None)
            if app is None:
                continue
            try:
                app.Quit()
            except Exception:
                pass
            setattr(self, key, None)
            self.counts[key] = 0
        gc.collect()

    # =========================================================
    # Word
    # =========================================================
    def _get_word(self):
        if self.word is None:
            self.word = win32com.client.DispatchEx("Word.Application")
            try:
                self.word.Visible = False
            except Exception:
                pass
            try:
                self.word.DisplayAlerts = 0  # wdAlertsNone
            except Exception:
                pass
        return self.word

    def extract_word(self, path: str) -> List[Segment]:
        word = self._get_word()
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
            self.counts["word"] += 1
            if self.counts["word"] >= self.recycle_every:
                self._recycle("word")
        return segments

    # =========================================================
    # PowerPoint
    # =========================================================
    def _get_ppt(self):
        if self.ppt is None:
            self.ppt = win32com.client.DispatchEx("PowerPoint.Application")
            try:
                self.ppt.DisplayAlerts = 1  # ppAlertsNone
            except Exception:
                pass
        return self.ppt

    def extract_powerpoint(self, path: str) -> List[Segment]:
        app = self._get_ppt()
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
            self.counts["ppt"] += 1
            if self.counts["ppt"] >= self.recycle_every:
                self._recycle("ppt")
        return segments

    # =========================================================
    # Excel (旧形式 .xls)
    # =========================================================
    def _get_excel(self):
        if self.excel is None:
            self.excel = win32com.client.DispatchEx("Excel.Application")
            try:
                self.excel.Visible = False
            except Exception:
                pass
            try:
                self.excel.DisplayAlerts = False
            except Exception:
                pass
            try:
                self.excel.AskToUpdateLinks = False
            except Exception:
                pass
        return self.excel

    def extract_excel_old(self, path: str) -> List[Segment]:
        excel = self._get_excel()
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
            self.counts["excel"] += 1
            if self.counts["excel"] >= self.recycle_every:
                self._recycle("excel")
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
