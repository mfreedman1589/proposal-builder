"""What CI runs before the sweep: build the stand-in decks (deck_skeleton.py)."""
import runpy
import sys
from pathlib import Path

sys.argv = ["deck_skeleton.py", "build"]
runpy.run_path(str(Path(__file__).with_name("deck_skeleton.py")), run_name="__main__")
