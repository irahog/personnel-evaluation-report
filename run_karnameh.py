"""PyInstaller entry script (a package's __main__ cannot be the entry point directly)."""
import sys

from karnameh.__main__ import main

if __name__ == "__main__":
    sys.exit(main())
