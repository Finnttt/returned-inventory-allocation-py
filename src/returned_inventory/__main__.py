"""``python -m returned_inventory`` 入口（转发到 cli.main）。"""

from .cli import main

raise SystemExit(main())
