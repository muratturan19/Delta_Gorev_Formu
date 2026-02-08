"""Command line entry point for running the Flask app.

This module makes it easy to run the application with ``python -m web_app``
which is handy for deployment platforms (e.g. Render) that expect the
process to bind to ``0.0.0.0`` and respect the ``PORT`` environment
variable.
"""
from __future__ import annotations

import os
from typing import Any

from . import create_app


def _get_port(default: int = 5002) -> int:
    raw_port = os.environ.get("PORT", str(default))
    try:
        return int(raw_port)
    except (TypeError, ValueError):
        return default


def _get_debug_flag(default: bool = False) -> bool:
    raw_value: Any = os.environ.get("FLASK_DEBUG") or os.environ.get("DEBUG")
    if raw_value is None:
        return default
    return str(raw_value).strip().lower() in {"1", "true", "yes", "on"}


def main() -> None:
    app = create_app()
    host = os.environ.get("HOST", "0.0.0.0")
    port = _get_port()
    debug = _get_debug_flag()
    app.run(host=host, port=port, debug=debug)


if __name__ == "__main__":
    main()
