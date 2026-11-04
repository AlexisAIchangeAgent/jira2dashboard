"""Where things live in the repo (the package is run from the repo, not installed)."""
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
SRC_DIR = PACKAGE_DIR.parent
REPO_ROOT = SRC_DIR.parent
VBA_DIR = REPO_ROOT / "vba"
PARAMS_DIR = REPO_ROOT / "params"
OUT_DIR = REPO_ROOT / "out"
CHILDREN_DIR = PACKAGE_DIR / "testbench" / "children"
