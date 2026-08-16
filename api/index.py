import sys
from pathlib import Path

# Ensure project root is on the path so `app.py`, templates, and static resolve.
ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import app  # noqa: E402
