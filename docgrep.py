#!/usr/bin/env python3
"""docgrep エントリポイント。

使い方:
    python docgrep.py "キーワード" [--mode keyword|regex|fuzzy] [--config config.yaml] ...

詳細は `python docgrep.py --help` を参照。
"""
import sys

from cli import main


if __name__ == "__main__":
    sys.exit(main())
