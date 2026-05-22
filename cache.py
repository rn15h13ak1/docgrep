"""抽出結果の SQLite インデックスキャッシュ。

ファイル path + mtime + size をキーに「抽出済み Segment 列」を保存し、
次回実行時にファイルメタが不変ならキャッシュからロードして抽出処理を
丸ごとスキップする。2 回目以降の検索が劇的に速くなる。

DB スキーマ:
  files     (path TEXT PRIMARY KEY, mtime REAL, size INTEGER,
             extractor TEXT, updated_at TEXT)
  segments  (path TEXT, idx INTEGER, locator TEXT, text TEXT,
             PRIMARY KEY (path, idx), FOREIGN KEY (path) → files)

キャッシュ自体は標準ライブラリ sqlite3 で実装するため追加インストール不要。
"""
from __future__ import annotations

import os
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

from search import Segment


_SCHEMA = """
CREATE TABLE IF NOT EXISTS files (
    path       TEXT PRIMARY KEY,
    mtime      REAL    NOT NULL,
    size       INTEGER NOT NULL,
    extractor  TEXT,
    updated_at TEXT
);
CREATE TABLE IF NOT EXISTS segments (
    path    TEXT    NOT NULL,
    idx     INTEGER NOT NULL,
    locator TEXT    NOT NULL,
    text    TEXT    NOT NULL,
    PRIMARY KEY (path, idx)
);
CREATE INDEX IF NOT EXISTS idx_segments_path ON segments(path);
"""


class SegmentCache:
    """抽出済み Segment を SQLite に永続化するキャッシュ。

    マルチスレッド書き込みに対応するため、書き込みは内部 lock で直列化する
    （sqlite3 自体も serialized モードを使う）。読み取りは並行可能。
    """

    def __init__(self, db_path: Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        # check_same_thread=False で別スレッドからも使えるようにする
        self._conn = sqlite3.connect(
            str(self.db_path),
            check_same_thread=False,
            isolation_level=None,  # autocommit
        )
        self._conn.execute("PRAGMA journal_mode=WAL;")
        self._conn.execute("PRAGMA synchronous=NORMAL;")
        self._conn.executescript(_SCHEMA)
        self._lock = threading.Lock()
        self.hits = 0
        self.misses = 0
        self.writes = 0

    # --- 読み取り ---
    def get(self, path: str) -> Optional[List[Segment]]:
        """キャッシュにヒットすれば Segment リストを返す。

        ファイル mtime/size が DB と一致しない場合はミス扱い。
        """
        try:
            stat = os.stat(path)
        except OSError:
            return None
        cur = self._conn.execute(
            "SELECT mtime, size FROM files WHERE path = ?", (path,)
        )
        row = cur.fetchone()
        if not row:
            self.misses += 1
            return None
        cached_mtime, cached_size = row
        if abs(cached_mtime - stat.st_mtime) > 1e-6 or cached_size != stat.st_size:
            # 内容が変わっているのでミス扱い（古いエントリは put 時に上書き）
            self.misses += 1
            return None
        rows = self._conn.execute(
            "SELECT locator, text FROM segments WHERE path = ? ORDER BY idx",
            (path,),
        ).fetchall()
        self.hits += 1
        return [Segment(text=t, locator=l) for (l, t) in rows]

    # --- 書き込み ---
    def put(self, path: str, segments: Iterable[Segment], extractor: str = "") -> None:
        """抽出結果を DB に書き込む。既存エントリは置き換える。"""
        try:
            stat = os.stat(path)
        except OSError:
            return
        seg_rows: List[Tuple[str, int, str, str]] = [
            (path, i, s.locator, s.text) for i, s in enumerate(segments)
        ]
        now = datetime.now().isoformat(timespec="seconds")
        with self._lock:
            cur = self._conn.cursor()
            try:
                cur.execute("BEGIN")
                cur.execute("DELETE FROM segments WHERE path = ?", (path,))
                cur.execute(
                    "INSERT OR REPLACE INTO files "
                    "(path, mtime, size, extractor, updated_at) "
                    "VALUES (?, ?, ?, ?, ?)",
                    (path, stat.st_mtime, stat.st_size, extractor, now),
                )
                if seg_rows:
                    cur.executemany(
                        "INSERT INTO segments (path, idx, locator, text) "
                        "VALUES (?, ?, ?, ?)",
                        seg_rows,
                    )
                cur.execute("COMMIT")
            except Exception:
                cur.execute("ROLLBACK")
                raise
        self.writes += 1

    def forget(self, path: str) -> None:
        """指定パスのキャッシュを削除する。"""
        with self._lock:
            cur = self._conn.cursor()
            try:
                cur.execute("BEGIN")
                cur.execute("DELETE FROM segments WHERE path = ?", (path,))
                cur.execute("DELETE FROM files    WHERE path = ?", (path,))
                cur.execute("COMMIT")
            except Exception:
                cur.execute("ROLLBACK")
                raise

    def vacuum(self) -> None:
        """孤児セグメントの掃除（オプション）。"""
        with self._lock:
            self._conn.execute(
                "DELETE FROM segments WHERE path NOT IN (SELECT path FROM files)"
            )
            self._conn.execute("VACUUM")

    def close(self) -> None:
        try:
            self._conn.close()
        except Exception:
            pass

    def __enter__(self) -> "SegmentCache":
        return self

    def __exit__(self, *exc) -> None:
        self.close()
