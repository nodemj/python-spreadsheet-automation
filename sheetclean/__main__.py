"""Allow running sheetclean via `python -m sheetclean`."""

import sys

from sheetclean.cli import main

if __name__ == "__main__":
    sys.exit(main())
