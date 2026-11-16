"""Repository-relative locations shared by the build, test and render scripts.

Layout: <repo>/build_tools (this folder), <repo>/VBA_modules, <repo>/MRI_Files_<year>.
Override the MRI source folder with the environment variable BUDGET_SRC.
"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))           # build_tools/
REPO = os.path.dirname(HERE)
VBA_DIR = os.path.join(REPO, "VBA_modules")
MRI_DIR = os.environ.get("BUDGET_SRC") or os.path.join(REPO, "MRI_Files_2027")
COLORS_JSON = os.path.join(HERE, "colors_detail.json")
