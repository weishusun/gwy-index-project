import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1] / "src"))

from iaci_index.crawling.step1_school_list import run_step1
from iaci_index.crawling.step2_official_site_search import run_step2
from iaci_index.crawling.step2_extra_info_urls import run_step2_extra_info
from iaci_index.crawling.step3_metrics_crawler import run_step3
from iaci_index.enrichment.step4_llm_metrics_completion import run_step4
from iaci_index.features.language import build_lri_features
from iaci_index.features.asean import build_arii_features
from iaci_index.features.text_intl import build_tli_features
from iaci_index.modeling.pca_model import run_pca_for_intl_index
from iaci_index.modeling.iaci_composite import compute_iaci_4d, prettify_scores


if __name__ == "__main__":
    run_step1()
    run_step2()
    run_step2_extra_info()
    run_step3()
    run_step4()
    build_lri_features()
    build_arii_features()
    build_tli_features()
    run_pca_for_intl_index()
    compute_iaci_4d()
    prettify_scores()
