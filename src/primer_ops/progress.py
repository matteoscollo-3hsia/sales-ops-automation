from __future__ import annotations

import itertools
import sys
import threading
import time
from contextlib import contextmanager


def format_seconds(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f}s"
    mins = int(seconds // 60)
    secs = seconds - mins * 60
    return f"{mins}m {secs:.0f}s"


@contextmanager
def spinner(message: str, interval_s: float = 0.1):
    stop = threading.Event()

    def run() -> None:
        for ch in itertools.cycle("|/-\\"):
            if stop.is_set():
                break
            sys.stdout.write(f"\r{message} {ch}")
            sys.stdout.flush()
            time.sleep(interval_s)

        # clear line
        sys.stdout.write("\r" + (" " * (len(message) + 2)) + "\r")
        sys.stdout.flush()

    t = threading.Thread(target=run, daemon=True)
    t.start()
    try:
        yield
    finally:
        stop.set()
        t.join(timeout=1)
