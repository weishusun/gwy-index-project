import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1] / "src"))

from iaci_index.enrichment.step4_llm_metrics_completion import run_step4

if __name__ == "__main__":
    run_step4()
