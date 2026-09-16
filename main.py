from __future__ import annotations

import sys

from worker_time_list import __version__


def main() -> int:
    if "--version" in sys.argv:
        print(__version__)
        return 0
    from worker_time_list.app import run
    run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
