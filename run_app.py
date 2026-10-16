"""PyInstaller entry script (a package's __main__ cannot be the entry point directly)."""
import sys

from personnel_evaluation_report.__main__ import main

if __name__ == "__main__":
    sys.exit(main())
