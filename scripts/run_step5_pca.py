import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1] / "src"))

from iaci_index.modeling.pca_model import run_pca_for_intl_index

if __name__ == "__main__":
    run_pca_for_intl_index()
