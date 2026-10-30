"""Allow `python -m xlvbesync ...`."""

from xlvbesync.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
