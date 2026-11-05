"""Where things live in the repo (the package is run from the repo, not installed)."""
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
SRC_DIR = PACKAGE_DIR.parent
REPO_ROOT = SRC_DIR.parent
VBA_DIR = REPO_ROOT / "vba"
PARAMS_DIR = REPO_ROOT / "params"
OUT_DIR = REPO_ROOT / "out"
RELEASE_DIR = OUT_DIR / "release"       # what is delivered: the workbook of `build`, the release packages
INTER_DIR = OUT_DIR / "inter"           # everything else: test workbooks, copies, exports (emptied by clean)
CHILDREN_DIR = PACKAGE_DIR / "testbench" / "children"
