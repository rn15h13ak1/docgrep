"""プロセス優先度制御（フォアグラウンド作業への干渉を抑える）。"""
from __future__ import annotations

import os
import sys


def apply_priority(level: str) -> None:
    """level = normal | below_normal | idle"""
    try:
        import psutil  # type: ignore
    except ImportError:
        return
    level = (level or "normal").lower()
    try:
        p = psutil.Process(os.getpid())
    except Exception:
        return

    try:
        if sys.platform == "win32":
            mapping = {
                "normal": psutil.NORMAL_PRIORITY_CLASS,
                "below_normal": psutil.BELOW_NORMAL_PRIORITY_CLASS,
                "idle": psutil.IDLE_PRIORITY_CLASS,
            }
            p.nice(mapping.get(level, psutil.NORMAL_PRIORITY_CLASS))
        else:
            mapping = {"normal": 0, "below_normal": 10, "idle": 19}
            p.nice(mapping.get(level, 0))
    except (psutil.AccessDenied, PermissionError):
        # 権限がない場合は黙って諦める（実行は継続）
        pass
    except Exception:
        pass
