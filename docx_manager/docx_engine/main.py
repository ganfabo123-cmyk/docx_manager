"""
main.py — Entry point for docx_engine_V3

Usage:
    python main.py [--host HOST] [--port PORT] [--debug]

Defaults:
    --host  0.0.0.0
    --port  5000
"""

import argparse
import logging
import sys
from pathlib import Path

# ── Startup checks ─────────────────────────────────────────────────────────────

_BASE = Path(__file__).parent

_REQUIRED = {
    "template":     _BASE / "templates" / "hit-template",
    "hit_config":   _BASE / "sections_config" / "hit_config.json",
    "extraction":   _BASE / "data" / "extraction.json",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)


def _check_prerequisites() -> bool:
    ok = True
    for name, path in _REQUIRED.items():
        if not path.exists():
            log.error("Missing required file/dir [%s]: %s", name, path)
            ok = False
    return ok


def main() -> None:
    parser = argparse.ArgumentParser(
        description="docx_engine_V3 — DOCX conversion server",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--host",  default="0.0.0.0", help="Bind address")
    parser.add_argument("--port",  default=5001, type=int, help="Port")
    parser.add_argument("--debug", action="store_true",    help="Flask debug mode")
    args = parser.parse_args()

    if not _check_prerequisites():
        log.error("Aborting — fix the missing files above and try again.")
        sys.exit(1)

    log.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    log.info("  docx_engine_V3")
    log.info("  http://%s:%d", args.host, args.port)
    log.info("  Template : %s", _REQUIRED["template"])
    log.info("  Config   : %s", _REQUIRED["hit_config"])
    log.info("  Outputs  : %s", _BASE / "outputs")
    log.info("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

    from server import app
    app.run(host=args.host, port=args.port, debug=args.debug)


if __name__ == "__main__":
    main()
