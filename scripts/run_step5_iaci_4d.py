import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1] / "src"))

from iaci_index.modeling.iaci_composite import compute_iaci_4d, prettify_scores

if __name__ == "__main__":
    compute_iaci_4d()
    prettify_scores()
