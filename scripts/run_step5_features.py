import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1] / "src"))

from iaci_index.features.language import build_lri_features
from iaci_index.features.asean import build_arii_features
from iaci_index.features.text_intl import build_tli_features


if __name__ == "__main__":
    build_lri_features()
    build_arii_features()
    build_tli_features()
