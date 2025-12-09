import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1] / "src"))

from iaci_index.crawling.step1_school_list import run_step1

if __name__ == "__main__":
    run_step1()
