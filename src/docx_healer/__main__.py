"""Allow `python -m docx_healer` to invoke the CLI."""

import sys

from .cli import main

if __name__ == "__main__":
    sys.exit(main())
