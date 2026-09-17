from __future__ import annotations

import sys

from worker_time_list import __version__


def main() -> int:
    if "--version" in sys.argv:
        print(__version__)
        return 0
    if "--smoke-gui" in sys.argv:
        from worker_time_list.app import WorkerTimeApp

        app = WorkerTimeApp()
        app.root.update_idletasks()
        app.root.destroy()
        return 0
    from worker_time_list.app import run

    run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
